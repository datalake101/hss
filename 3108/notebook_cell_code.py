# ==============================================================================
# Health Sciences Premium Presentation Deck Generator (Jupyter Notebook Ready)
# Converts structured lecture .txt files into premium 16:9 PowerPoint decks.
# ==============================================================================

import glob
import math
import os
import re
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- design system

PALETTES = {
    "crimson": {  # deep crimson + charcoal body + antique gold (default)
        "panel": "952210",   # requested: 952210 for filled box
        "accent": "C9A45C",  # antique gold accent line & ribbons
        "label": "F0D5D2",   # soft rose tint on crimson panel
        "title": "FFFFFF",   # crisp white title on crimson panel
        "body": "373D42",    # requested: 373D42 for text
        "lead_in": "952210", # matching crimson for bold lead-in tags
        "footer": "99A1AA",  # subtle slate footer
    },
    "midnight": {  # deep navy + antique gold
        "panel": "102A43", "accent": "C9A45C", "label": "8FA8C2",
        "title": "FFFFFF", "body": "373D42", "lead_in": "102A43", "footer": "99A1AA",
    },
    "forest": {    # deep emerald + champagne
        "panel": "1B3A32", "accent": "C7A86B", "label": "9CB8AC",
        "title": "FFFFFF", "body": "373D42", "lead_in": "1B3A32", "footer": "9AA5A0",
    },
    "graphite": {  # charcoal + teal
        "panel": "262A2E", "accent": "3EA896", "label": "A8AEB5",
        "title": "FFFFFF", "body": "373D42", "lead_in": "262A2E", "footer": "9BA1A6",
    },
    "wine": {      # oxblood + soft gold
        "panel": "3D1F28", "accent": "C9A26B", "label": "C0A2A9",
        "title": "FFFFFF", "body": "373D42", "lead_in": "3D1F28", "footer": "A79CA0",
    },
}

TITLE_FONT   = "Lovelo"       # geometric modern display font for titles
BODY_FONT    = "Segoe UI"     # clean Segoe font for main text
BULLET_EMOJI = "🌱"          # seedling bullet 🌱

# geometry (inches) — 16:9 canvas
SLIDE_W, SLIDE_H = 13.333, 7.5
PANEL_W          = 4.10
EDGE_W           = 0.055    # gold ribbon on panel's right edge
BODY_X, BODY_Y   = 4.80, 0.72
BODY_W, BODY_H   = 8.00, 5.95

# ---------------------------------------------------------------- parsing

def clean_em_dashes(text):
    text = re.sub(r"([\")])\s*[—–]\s*", r"\1: ", text)
    text = re.sub(r"\s*[—–]\s*", ", ", text)
    text = re.sub(r",\s*,", ",", text)
    text = re.sub(r":\s*,", ": ", text)
    text = text.replace("—", "").replace("–", "")
    return text

def parse(path):
    text = open(path, encoding="utf-8").read()
    text = clean_em_dashes(text)

    deck_title = "Lecture"
    m = re.search(r"LECTURE\s+\d+[^\S\r\n]*[:\-]\s*([^\r\n]+)", text[:500], re.I)
    if m:
        deck_title = m.group(1).strip()

    course = None
    m = re.search(r"Course:\s*(.+?)(?:,\s*Lecture|\n|$)", text, re.I)
    if m:
        course = m.group(1).strip()
    if not course:
        course = "Foundations of Public Health Practice"

    m_lec = re.search(r"LECTURE\s+(\d+)", text[:500], re.I)
    if m_lec:
        label = f"{course} \u2022 Lecture {m_lec.group(1)}"
    else:
        label = course

    # Split text into slide blocks (supports [SLIDE 1], SLIDE 1: ..., etc.)
    blocks = [
        b.strip()
        for b in re.split(r"(?:\r?\n)+\s*(?=\[?SLIDE\s+\d+)", text, flags=re.I)
        if b.strip()
    ]

    slides = []
    for block in blocks:
        lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
        if not lines:
            continue

        m_slide = re.match(
            r"^\s*\[?SLIDE\s+(\d+)\]?(?:\s*[:\-–—,]\s*(.*))?$",
            lines[0],
            re.I,
        )
        if not m_slide:
            continue

        num = int(m_slide.group(1))
        inline_title = (m_slide.group(2) or "").strip()
        inline_title = re.sub(r"^[A-Z]+\d+:\s*", "", inline_title)
        inline_title = re.sub(r"^\[[A-Z\s]+\]\s*", "", inline_title)

        body_lines = lines[1:]

        title = inline_title
        has_content_marker = False
        content_lines = []

        for line_str in body_lines:
            m_title = re.match(r"^TITLE:\s*(.*)$", line_str, re.I)
            if m_title:
                title = m_title.group(1).strip()
                continue
            if re.match(r"^CONTENT:\s*$", line_str, re.I):
                has_content_marker = True
                continue
            content_lines.append(line_str)

        if (
            (not title or title.upper() in ("TITLE SLIDE", "TITLE", "TITLE SLIDE:"))
            and deck_title
            and deck_title != "Lecture"
        ):
            title = deck_title
        elif not title and content_lines:
            title = content_lines.pop(0)

        bullets = []
        if has_content_marker:
            # Legacy format with explicit CONTENT: marker
            bullets = [
                re.sub(r"^[-*•]\s*", "", ln).strip()
                for ln in content_lines
                if re.match(r"^[-*•]", ln)
            ]
        else:
            # Modern format
            for ln in content_lines:
                cleaned = re.sub(r"^[-*•]\s*", "", ln).strip()
                if cleaned and cleaned not in (":", "-"):
                    bullets.append(cleaned)

        slides.append({
            "num": num,
            "title": title,
            "bullets": bullets,
        })

    if (not deck_title or deck_title == "Lecture") and slides and slides[0].get("title"):
        deck_title = slides[0]["title"]

    return deck_title, label, slides

# ---------------------------------------------------------------- helpers

def rect(slide, x, y, w, h, hexcolor):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                               Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = RGBColor.from_string(hexcolor)
    s.line.fill.background()
    s.shadow.inherit = False
    return s

def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf

def style_run(run, text, font, size, hexcolor, bold=False, tracking=None):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(hexcolor)
    if tracking:
        run._r.get_or_add_rPr().set("spc", str(tracking))

def set_hanging_indent(paragraph, indent_in=0.38):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(Inches(indent_in)))
    pPr.set("indent", str(-Inches(indent_in)))
    buNone = pPr.makeelement(qn("a:buNone"), {})
    pPr.append(buNone)

LEAD_RE = re.compile(r"^((?:\d+\.\s*)?[\"']?[A-Z0-9~](?:(?!\.\s+[A-Z])[^:\n!?]){0,60}):\s+(.+)$", re.S)

def fit_body_size(bullets, width_in, height_in, max_size=18.0, min_size=10.5):
    size = max_size
    while size > min_size:
        cpl = max(20, int((width_in - 0.45) * 72 / (0.50 * size)))
        total = sum(max(1, math.ceil((len(b) + 4) / cpl)) * size * 1.22 / 72
                    for b in bullets)
        spacing_ratio = 0.35 if len(bullets) > 12 else 0.58
        total += max(0, len(bullets) - 1) * size * spacing_ratio / 72
        if total <= height_in:
            break
        size -= 0.5
    return size

def fit_title_size(title, width_in, max_size=28, min_size=18):
    size = max_size
    while size > min_size:
        cpl = max(8, int(width_in * 72 / (0.58 * size)))
        if math.ceil(len(title) / cpl) * size * 1.18 / 72 <= 4.8:
            break
        size -= 1
    return size


# ---------------------------------------------------------------- deck builder

def find_uottawa_logo():
    candidates = [
        "uOttawa logo",
        "uOttawa logo.png",
        "uOttawa logo.jpg",
        "uottawa logo.png",
        "uottawa_logo.png",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    if "__file__" in globals():
        base_dir = os.path.dirname(os.path.abspath(__file__))
        for c in candidates:
            p = os.path.join(base_dir, c)
            if os.path.exists(p):
                return p
    matches = glob.glob("*uOttawa*logo*") + glob.glob("*uottawa*logo*")
    return matches[0] if matches else None

def build(deck_title, label, slides, palette_name, out_path):
    C = PALETTES.get(palette_name, PALETTES["crimson"])
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
    blank = prs.slide_layouts[6]
    total = len(slides)
    logo_file = find_uottawa_logo()

    for i, s in enumerate(slides, start=1):
        slide = prs.slides.add_slide(blank)

        # ---- left panel (no orange edge line)
        rect(slide, 0, 0, PANEL_W, SLIDE_H, C["panel"])

        # ---- title centered in left panel (horizontal bar removed)
        tf = textbox(slide, 0.45, 0.80, 3.20, 5.80, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.line_spacing = 1.06
        style_run(p.add_run(), s["title"], TITLE_FONT,
                  fit_title_size(s["title"], 3.20), C["title"], bold=True)

        # ---- slide counter (all slides except first and last, formatted as 2/32)
        if 1 < i < total:
            tf = textbox(slide, 0.45, 6.85, 3.20, 0.35)
            p_num = tf.paragraphs[0]
            p_num.alignment = PP_ALIGN.CENTER
            style_run(p_num.add_run(), f"{i}/{total}",
                      BODY_FONT, 10, C["label"], tracking=150)


        # ---- body bullets (vertically centered in white canvas)
        if s["bullets"]:
            size = fit_body_size(s["bullets"], BODY_W, BODY_H)
            tf = textbox(slide, BODY_X, BODY_Y, BODY_W, BODY_H, anchor=MSO_ANCHOR.MIDDLE)


            for j, bullet in enumerate(s["bullets"]):
                p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                p.line_spacing = 1.14
                spacing_ratio = 0.35 if len(s["bullets"]) > 12 else 0.58
                p.space_after = Pt(size * spacing_ratio)
                set_hanging_indent(p, indent_in=0.38)

                # Seedling bullet 🌱
                style_run(p.add_run(), f"{BULLET_EMOJI}  ", BODY_FONT, size, C["body"])

                m = LEAD_RE.match(bullet)
                if m:
                    style_run(p.add_run(), m.group(1) + ":  ",
                              BODY_FONT, size, C["lead_in"], bold=True)
                    style_run(p.add_run(), m.group(2), BODY_FONT, size, C["body"])
                else:
                    style_run(p.add_run(), bullet, BODY_FONT, size, C["body"])

        # ---- uOttawa logo at bottom right corner (all slides except first and last)
        # (Footer text removed)
        if 1 < i < total and logo_file:
            logo_h = 0.55
            logo_w = logo_h * 1867 / 1570
            logo_x = 12.80 - logo_w
            logo_y = 6.65
            slide.shapes.add_picture(logo_file, Inches(logo_x), Inches(logo_y),
                                     width=Inches(logo_w), height=Inches(logo_h))

    prs.core_properties.title = deck_title
    prs.save(out_path)
    print(f" Saved: {out_path} ({total} slides | palette: {palette_name})")

# ---------------------------------------------------------------- execution
# In Jupyter Notebook, we run this directly without argparse.
# Change palette to: 'crimson', 'midnight', 'forest', 'graphite', or 'wine'
SELECTED_PALETTE = "crimson"

# Find all lecture text files in the current folder and convert them:
raw_candidates = glob.glob("[lL]ecture*.txt")
lecture_files = sorted(list(dict.fromkeys([f for f in raw_candidates if os.path.isfile(f)])))

if not lecture_files:
    print("⚠️ No '[lL]ecture*.txt' files found in current directory:", os.getcwd())
else:
    print(f"🚀 Found {len(lecture_files)} lecture files. Generating decks...\n")
    for txt_file in lecture_files:
        output_pptx = re.sub(r"\.txt$", "", txt_file) + "_premium.pptx"
        deck_title, label, slides = parse(txt_file)
        build(deck_title, label, slides, SELECTED_PALETTE, output_pptx)

    print("\n🎉 Done! All PowerPoint decks generated successfully!")
