"""
Splunk SPL Handbook — HTML to PowerPoint Converter + OneDrive Uploader
======================================================================
Converts index.html (Splunk SPL Threat Hunter's Field Reference) into a
professional .pptx presentation and optionally uploads it to OneDrive.

Requirements:
    pip install python-pptx beautifulsoup4 lxml requests msal

Usage:
    # Convert only (generates .pptx locally)
    python html_to_pptx.py

    # Convert + Upload to OneDrive
    python html_to_pptx.py --upload

    # With custom OneDrive folder path
    python html_to_pptx.py --upload --onedrive-folder "Documents/Presentations"

OneDrive Setup (for upload):
    1. Register an app at https://portal.azure.com → Azure Active Directory → App registrations
    2. Add redirect URI: http://localhost
    3. API Permissions: Files.ReadWrite (delegated)
    4. Set environment variables:
         export AZURE_CLIENT_ID="your-client-id"
         export AZURE_TENANT_ID="your-tenant-id"
       Or create a .env file with these values.

Author: avx-sdeepak
Date: 2026-02-25
"""

import os
import sys
import argparse
import re
from pathlib import Path
from io import BytesIO

# ── Third-party imports ──────────────────────────────────────────────
try:
    from bs4 import BeautifulSoup, NavigableString
except ImportError:
    sys.exit("ERROR: Missing 'beautifulsoup4'. Install with: pip install beautifulsoup4 lxml")

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
except ImportError:
    sys.exit("ERROR: Missing 'python-pptx'. Install with: pip install python-pptx")


# ══════════════════════════════════════════════════════════════════════
# THEME / COLOR PALETTE (matches the dark GitHub-style HTML theme)
# ══════════════════════════════════════════════════════════════════════
class Theme:
    BG_DARK      = RGBColor(0x0D, 0x11, 0x17)   # --bg
    BG_SURFACE   = RGBColor(0x16, 0x1B, 0x22)   # --surface
    BORDER       = RGBColor(0x30, 0x36, 0x3D)   # --border
    TEXT         = RGBColor(0xE6, 0xED, 0xF3)   # --text
    TEXT_MUTED   = RGBColor(0x8B, 0x94, 0x9E)   # --text-muted
    ACCENT_BLUE  = RGBColor(0x58, 0xA6, 0xFF)   # --accent (section headings)
    ACCENT_GREEN = RGBColor(0x3F, 0xB9, 0x50)   # --accent2 (h3 / tips)
    ACCENT_ORANGE= RGBColor(0xF0, 0x88, 0x3E)   # --accent3 (h4 / inline code)
    ACCENT_PINK  = RGBColor(0xF7, 0x78, 0xBA)   # --accent4 (hunt card titles)
    CODE_BG      = RGBColor(0x1C, 0x21, 0x29)   # --code-bg
    RED          = RGBColor(0xF8, 0x51, 0x49)   # --red
    YELLOW       = RGBColor(0xD2, 0x99, 0x22)   # --yellow
    WHITE        = RGBColor(0xFF, 0xFF, 0xFF)


# ══════════════════════════════════════════════════════════════════════
# SLIDE DIMENSIONS — Widescreen 16:9
# ══════════════════════════════════════════════════════════════════════
SLIDE_WIDTH  = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


# ══════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════════════

def set_slide_bg(slide, color):
    """Set solid background color for a slide."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, left, top, width, height):
    """Add a textbox and return its text_frame."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    return tf


def add_run(paragraph, text, font_size=14, color=Theme.TEXT, bold=False, italic=False, font_name="Calibri"):
    """Add a styled run to a paragraph."""
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font_name
    return run


def add_code_run(paragraph, text, font_size=11, color=Theme.TEXT):
    """Add a monospace code run."""
    return add_run(paragraph, text, font_size=font_size, color=color, font_name="Consolas")


def strip_html_tags(html_string):
    """Remove HTML tags and return plain text."""
    soup = BeautifulSoup(html_string, "lxml")
    return soup.get_text()


def get_text_from_element(element):
    """Recursively extract text from a BS4 element."""
    if isinstance(element, NavigableString):
        return str(element)
    texts = []
    for child in element.children:
        texts.append(get_text_from_element(child))
    return "".join(texts)


def chunk_text(text, max_lines=18):
    """Split text into chunks of max_lines lines each."""
    lines = text.split("\n")
    chunks = []
    for i in range(0, len(lines), max_lines):
        chunks.append("\n".join(lines[i:i + max_lines]))
    return chunks


def add_decorative_bar(slide, left, top, width, height, color):
    """Add a colored decorative rectangle bar."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


# ══════════════════════════════════════════════════════════════════════
# SLIDE GENERATORS
# ══════════════════════════════════════════════════════════════════════

def create_title_slide(prs, title, subtitle, badge_text=""):
    """Create the hero / title slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    set_slide_bg(slide, Theme.BG_DARK)

    # Decorative top bar
    add_decorative_bar(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.06), Theme.ACCENT_BLUE)

    # Title
    tf = add_textbox(slide, Inches(1), Inches(1.8), Inches(11.3), Inches(1.5))
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_run(p, title, font_size=44, color=Theme.ACCENT_BLUE, bold=True)

    # Subtitle
    tf2 = add_textbox(slide, Inches(2), Inches(3.5), Inches(9.3), Inches(1))
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    add_run(p2, subtitle, font_size=18, color=Theme.TEXT_MUTED)

    # Badge
    if badge_text:
        tf3 = add_textbox(slide, Inches(4), Inches(4.8), Inches(5.3), Inches(0.6))
        p3 = tf3.paragraphs[0]
        p3.alignment = PP_ALIGN.CENTER
        add_run(p3, f"  {badge_text}  ", font_size=13, color=Theme.ACCENT_ORANGE, italic=True)

    # Bottom bar
    add_decorative_bar(slide, Inches(0), Inches(7.35), SLIDE_WIDTH, Inches(0.15), Theme.ACCENT_PINK)

    return slide


def create_toc_slide(prs, sections):
    """Create a Table of Contents slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, Theme.BG_DARK)

    add_decorative_bar(slide, Inches(0), Inches(0), Inches(0.08), SLIDE_HEIGHT, Theme.ACCENT_BLUE)

    # Title
    tf = add_textbox(slide, Inches(0.6), Inches(0.3), Inches(12), Inches(0.8))
    p = tf.paragraphs[0]
    add_run(p, "TABLE OF CONTENTS", font_size=28, color=Theme.ACCENT_BLUE, bold=True)

    # Divider
    add_decorative_bar(slide, Inches(0.6), Inches(1.1), Inches(4), Inches(0.03), Theme.BORDER)

    # Two-column layout for TOC
    left_sections = sections[:7]
    right_sections = sections[7:]

    y_offset = 1.4
    for i, sec in enumerate(left_sections):
        tf_item = add_textbox(slide, Inches(0.8), Inches(y_offset), Inches(5.5), Inches(0.5))
        p_item = tf_item.paragraphs[0]
        add_run(p_item, f"{i+1:02d}  ", font_size=14, color=Theme.TEXT_MUTED)
        add_run(p_item, sec, font_size=14, color=Theme.TEXT, bold=False)
        y_offset += 0.55

    y_offset = 1.4
    for i, sec in enumerate(right_sections):
        tf_item = add_textbox(slide, Inches(7), Inches(y_offset), Inches(5.5), Inches(0.5))
        p_item = tf_item.paragraphs[0]
        add_run(p_item, f"{i+8:02d}  ", font_size=14, color=Theme.TEXT_MUTED)
        add_run(p_item, sec, font_size=14, color=Theme.TEXT, bold=False)
        y_offset += 0.55

    return slide


def create_section_title_slide(prs, section_num, section_title):
    """Create a section divider slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, Theme.BG_SURFACE)

    # Large section number
    tf_num = add_textbox(slide, Inches(1), Inches(1.5), Inches(3), Inches(2))
    p_num = tf_num.paragraphs[0]
    add_run(p_num, f"{section_num:02d}", font_size=72, color=Theme.BORDER, bold=True)

    # Section title
    tf_title = add_textbox(slide, Inches(1), Inches(3.5), Inches(11), Inches(1.5))
    p_title = tf_title.paragraphs[0]
    add_run(p_title, section_title, font_size=36, color=Theme.ACCENT_BLUE, bold=True)

    # Decorative line
    add_decorative_bar(slide, Inches(1), Inches(5.2), Inches(3), Inches(0.04), Theme.ACCENT_BLUE)

    return slide


def create_content_slide(prs, title, content_blocks, section_num=0):
    """
    Create a content slide with title and mixed content blocks.
    content_blocks: list of dicts with 'type' and 'content' keys.
        type: 'text', 'code', 'table', 'callout', 'hunt_card'
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, Theme.BG_DARK)

    # Top accent bar
    add_decorative_bar(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.04), Theme.ACCENT_BLUE)

    # Slide title
    tf_title = add_textbox(slide, Inches(0.5), Inches(0.2), Inches(12), Inches(0.7))
    p_title = tf_title.paragraphs[0]
    if section_num:
        add_run(p_title, f"{section_num}. ", font_size=10, color=Theme.TEXT_MUTED)
    add_run(p_title, title, font_size=22, color=Theme.ACCENT_GREEN, bold=True)

    # Underline
    add_decorative_bar(slide, Inches(0.5), Inches(0.85), Inches(12.3), Inches(0.02), Theme.BORDER)

    y_pos = 1.1
    max_y = 7.0

    for block in content_blocks:
        if y_pos >= max_y:
            break

        if block["type"] == "text":
            remaining_height = max_y - y_pos
            tf = add_textbox(slide, Inches(0.6), Inches(y_pos), Inches(12), Inches(min(remaining_height, 1.0)))
            p = tf.paragraphs[0]
            text = block["content"]
            # Handle bold fragments
            parts = re.split(r'(\*\*.*?\*\*)', text)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    add_run(p, part[2:-2], font_size=13, color=Theme.TEXT, bold=True)
                else:
                    add_run(p, part, font_size=13, color=Theme.TEXT_MUTED)
            y_pos += 0.5

        elif block["type"] == "code":
            code_text = block["content"]
            lines = code_text.strip().split("\n")
            num_lines = len(lines)
            box_height = min(max(num_lines * 0.22 + 0.3, 0.6), max_y - y_pos)

            if box_height < 0.4:
                break

            # Code background box
            code_shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(0.5), Inches(y_pos),
                Inches(12.3), Inches(box_height)
            )
            code_shape.fill.solid()
            code_shape.fill.fore_color.rgb = Theme.CODE_BG
            code_shape.line.color.rgb = Theme.BORDER
            code_shape.line.width = Pt(1)

            # Code text
            tf_code = add_textbox(slide, Inches(0.7), Inches(y_pos + 0.1),
                                   Inches(11.8), Inches(box_height - 0.2))
            p_code = tf_code.paragraphs[0]
            # Truncate if too many lines
            display_lines = lines[:int((box_height - 0.3) / 0.22)]
            truncated = "\n".join(display_lines)
            if len(display_lines) < len(lines):
                truncated += "\n  ... (continued)"
            add_code_run(p_code, truncated, font_size=10, color=Theme.TEXT)
            y_pos += box_height + 0.15

        elif block["type"] == "callout":
            remaining = max_y - y_pos
            box_h = min(0.7, remaining)
            if box_h < 0.3:
                break
            # Left border bar
            callout_color = Theme.YELLOW
            if "danger" in block.get("style", ""):
                callout_color = Theme.RED
            elif "tip" in block.get("style", ""):
                callout_color = Theme.ACCENT_GREEN

            add_decorative_bar(slide, Inches(0.5), Inches(y_pos), Inches(0.06), Inches(box_h), callout_color)

            tf_callout = add_textbox(slide, Inches(0.7), Inches(y_pos), Inches(11.8), Inches(box_h))
            p_callout = tf_callout.paragraphs[0]
            add_run(p_callout, block["content"], font_size=12, color=Theme.YELLOW, italic=True)
            y_pos += box_h + 0.15

        elif block["type"] == "table_data":
            rows = block["content"]  # list of lists
            if not rows:
                continue
            num_cols = len(rows[0])
            num_rows = min(len(rows), int((max_y - y_pos - 0.2) / 0.32))
            if num_rows < 2:
                break

            col_width = min(12.0 / num_cols, 4.0)
            table_width = col_width * num_cols
            row_height = 0.32

            table_shape = slide.shapes.add_table(
                num_rows, num_cols,
                Inches(0.6), Inches(y_pos),
                Inches(table_width), Inches(num_rows * row_height)
            )
            table = table_shape.table

            for r_idx in range(num_rows):
                for c_idx in range(num_cols):
                    cell = table.cell(r_idx, c_idx)
                    cell_text = rows[r_idx][c_idx] if c_idx < len(rows[r_idx]) else ""
                    cell.text = cell_text

                    # Style cell
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(10)
                        paragraph.font.name = "Calibri"
                        if r_idx == 0:
                            paragraph.font.bold = True
                            paragraph.font.color.rgb = Theme.ACCENT_BLUE
                        else:
                            paragraph.font.color.rgb = Theme.TEXT

                    # Cell fill
                    cell_fill = cell.fill
                    cell_fill.solid()
                    if r_idx == 0:
                        cell_fill.fore_color.rgb = Theme.BG_SURFACE
                    elif r_idx % 2 == 0:
                        cell_fill.fore_color.rgb = RGBColor(0x12, 0x16, 0x1C)
                    else:
                        cell_fill.fore_color.rgb = Theme.BG_DARK

            y_pos += num_rows * row_height + 0.2

    return slide


def create_hunt_card_slide(prs, title, tags, code_text, description=""):
    """Create a slide for a Threat Hunting query card."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, Theme.BG_DARK)

    # Top accent
    add_decorative_bar(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.05), Theme.ACCENT_PINK)

    # Hunt card title
    tf_title = add_textbox(slide, Inches(0.5), Inches(0.2), Inches(12), Inches(0.7))
    p_title = tf_title.paragraphs[0]
    add_run(p_title, "🎯 ", font_size=22, color=Theme.ACCENT_PINK)
    add_run(p_title, title, font_size=22, color=Theme.ACCENT_PINK, bold=True)

    # Tags
    y_pos = 0.9
    if tags:
        tf_tags = add_textbox(slide, Inches(0.6), Inches(y_pos), Inches(11), Inches(0.4))
        p_tags = tf_tags.paragraphs[0]
        for tag in tags:
            if "MITRE" in tag.upper() or tag.startswith("T"):
                add_run(p_tags, f" {tag} ", font_size=10, color=Theme.ACCENT_BLUE, bold=True)
                add_run(p_tags, "  ", font_size=10, color=Theme.TEXT)
            else:
                add_run(p_tags, f" {tag} ", font_size=10, color=Theme.ACCENT_GREEN, bold=True)
                add_run(p_tags, "  ", font_size=10, color=Theme.TEXT)
        y_pos += 0.45

    # Description
    if description:
        tf_desc = add_textbox(slide, Inches(0.6), Inches(y_pos), Inches(11), Inches(0.4))
        p_desc = tf_desc.paragraphs[0]
        add_run(p_desc, description, font_size=12, color=Theme.ACCENT_GREEN, italic=True)
        y_pos += 0.4

    # Code block
    code_lines = code_text.strip().split("\n")
    code_height = min(max(len(code_lines) * 0.24 + 0.3, 0.8), 7.0 - y_pos - 0.2)

    code_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.4), Inches(y_pos + 0.1),
        Inches(12.5), Inches(code_height)
    )
    code_shape.fill.solid()
    code_shape.fill.fore_color.rgb = Theme.CODE_BG
    code_shape.line.color.rgb = Theme.BORDER
    code_shape.line.width = Pt(1)

    tf_code = add_textbox(slide, Inches(0.6), Inches(y_pos + 0.2),
                           Inches(12), Inches(code_height - 0.2))
    p_code = tf_code.paragraphs[0]
    max_display = int((code_height - 0.3) / 0.24)
    display = code_lines[:max_display]
    if len(display) < len(code_lines):
        display.append("  ... (see handbook for full query)")
    add_code_run(p_code, "\n".join(display), font_size=10, color=Theme.TEXT)

    return slide


def create_closing_slide(prs):
    """Create a closing / thank-you slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, Theme.BG_DARK)

    add_decorative_bar(slide, Inches(0), Inches(3.4), SLIDE_WIDTH, Inches(0.04), Theme.ACCENT_BLUE)

    tf = add_textbox(slide, Inches(1), Inches(2), Inches(11.3), Inches(1.2))
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_run(p, "Splunk SPL — Threat Hunter's Field Reference", font_size=32, color=Theme.ACCENT_BLUE, bold=True)

    tf2 = add_textbox(slide, Inches(2), Inches(4), Inches(9.3), Inches(1))
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    add_run(p2, "From raw log parsing through statistical anomaly detection,\ncorrelation, and visualization.",
            font_size=16, color=Theme.TEXT_MUTED)

    tf3 = add_textbox(slide, Inches(3), Inches(5.5), Inches(7.3), Inches(0.5))
    p3 = tf3.paragraphs[0]
    p3.alignment = PP_ALIGN.CENTER
    add_run(p3, "github.com/avx-sdeepak/splunk-spl-handbook", font_size=12, color=Theme.TEXT_MUTED, italic=True)

    add_decorative_bar(slide, Inches(0), Inches(7.35), SLIDE_WIDTH, Inches(0.15), Theme.ACCENT_PINK)

    return slide


# ══════════════════════════════════════════════════════════════════════
# HTML PARSER → SLIDE DATA
# ══════════════════════════════════════════════════════════════════════

def parse_html(html_path):
    """Parse the HTML file and extract structured slide data."""
    with open(html_path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "lxml")

    # ── Extract hero ──
    hero_h1 = soup.find("h1")
    hero_title = hero_h1.get_text() if hero_h1 else "Splunk SPL Reference"

    hero_p = soup.find("div", class_="hero")
    hero_subtitle = ""
    hero_badge = ""
    if hero_p:
        p_tag = hero_p.find("p")
        if p_tag:
            hero_subtitle = p_tag.get_text()
        badge_tag = hero_p.find("span", class_="badge")
        if badge_tag:
            hero_badge = badge_tag.get_text()

    # ── Extract sections ──
    sections = []
    h2_tags = soup.find_all("h2")

    for h2 in h2_tags:
        section_title = h2.get_text().strip()
        # Remove leading number+dot if present
        clean_title = re.sub(r"^\d+\.\s*", "", section_title)

        # Collect all sibling elements until the next h2
        content_elements = []
        sibling = h2.find_next_sibling()
        while sibling and sibling.name != "h2":
            content_elements.append(sibling)
            sibling = sibling.find_next_sibling()

        # Parse content elements into blocks
        subsections = []
        current_sub = {"title": clean_title, "blocks": []}

        for elem in content_elements:
            if elem.name == "h3":
                if current_sub["blocks"]:
                    subsections.append(current_sub)
                current_sub = {"title": elem.get_text().strip(), "blocks": []}

            elif elem.name == "pre":
                code_text = elem.get_text()
                current_sub["blocks"].append({"type": "code", "content": code_text})

            elif elem.name == "table":
                rows = []
                for tr in elem.find_all("tr"):
                    cells = [td.get_text().strip() for td in tr.find_all(["th", "td"])]
                    if cells:
                        rows.append(cells)
                if rows:
                    current_sub["blocks"].append({"type": "table_data", "content": rows})

            elif elem.name == "p":
                text = elem.get_text().strip()
                if text:
                    current_sub["blocks"].append({"type": "text", "content": text})

            elif elem.name == "div" and "callout" in elem.get("class", []):
                text = elem.get_text().strip()
                classes = " ".join(elem.get("class", []))
                current_sub["blocks"].append({
                    "type": "callout",
                    "content": text,
                    "style": classes
                })

            elif elem.name == "div" and "hunt-card" in elem.get("class", []):
                card_title_tag = elem.find("h4")
                card_title = card_title_tag.get_text().strip() if card_title_tag else "Hunt Query"

                tags = [t.get_text().strip() for t in elem.find_all("span", class_="tag")]

                card_code = ""
                card_pre = elem.find("pre")
                if card_pre:
                    card_code = card_pre.get_text()

                card_desc = ""
                card_p = elem.find("p")
                if card_p:
                    card_desc = card_p.get_text().strip()

                current_sub["blocks"].append({
                    "type": "hunt_card",
                    "title": card_title,
                    "tags": tags,
                    "code": card_code,
                    "description": card_desc
                })

        if current_sub["blocks"]:
            subsections.append(current_sub)

        sections.append({
            "title": clean_title,
            "subsections": subsections
        })

    return {
        "hero_title": hero_title,
        "hero_subtitle": hero_subtitle,
        "hero_badge": hero_badge,
        "sections": sections
    }


# ══════════════════════════════════════════════════════════════════════
# MAIN CONVERSION
# ══════════════════════════════════════════════════════════════════════

def convert_html_to_pptx(html_path, output_path):
    """Convert HTML handbook to PowerPoint presentation."""
    print(f"📄 Parsing HTML: {html_path}")
    data = parse_html(html_path)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # ── 1. Title slide ──
    print("  → Creating title slide...")
    create_title_slide(prs, data["hero_title"], data["hero_subtitle"], data["hero_badge"])

    # ── 2. Table of Contents ──
    section_titles = [s["title"] for s in data["sections"]]
    print(f"  → Creating TOC ({len(section_titles)} sections)...")
    create_toc_slide(prs, section_titles)

    # ── 3. Section slides ──
    for sec_idx, section in enumerate(data["sections"], 1):
        print(f"  → Section {sec_idx}: {section['title']}")

        # Section title slide
        create_section_title_slide(prs, sec_idx, section["title"])

        # Content slides for each subsection
        for sub in section["subsections"]:
            # Check for hunt cards — they get their own slide
            hunt_cards = [b for b in sub["blocks"] if b["type"] == "hunt_card"]
            other_blocks = [b for b in sub["blocks"] if b["type"] != "hunt_card"]

            # Regular content slide
            if other_blocks:
                # Split into multiple slides if too much content
                code_blocks = [b for b in other_blocks if b["type"] == "code"]
                non_code = [b for b in other_blocks if b["type"] != "code"]

                # First slide: non-code + first code block
                first_slide_blocks = non_code[:]
                if code_blocks:
                    first_slide_blocks.append(code_blocks[0])
                if first_slide_blocks:
                    create_content_slide(prs, sub["title"], first_slide_blocks, section_num=sec_idx)

                # Additional slides for remaining code blocks
                for cb in code_blocks[1:]:
                    create_content_slide(prs, f"{sub['title']} (cont.)", [cb], section_num=sec_idx)

            # Hunt card slides
            for card in hunt_cards:
                create_hunt_card_slide(
                    prs,
                    card["title"],
                    card["tags"],
                    card["code"],
                    card.get("description", "")
                )

    # ── 4. Closing slide ──
    print("  → Creating closing slide...")
    create_closing_slide(prs)

    # ── Save ──
    prs.save(output_path)
    total_slides = len(prs.slides)
    file_size = os.path.getsize(output_path) / 1024
    print(f"\n✅ Presentation saved: {output_path}")
    print(f"   Slides: {total_slides}")
    print(f"   Size: {file_size:.1f} KB")

    return output_path


# ══════════════════════════════════════════════════════════════════════
# ONEDRIVE UPLOAD
# ══════════════════════════════════════════════════════════════════════

def upload_to_onedrive(file_path, folder_path="Documents"):
    """
    Upload the generated .pptx to OneDrive using Microsoft Graph API.

    Requires:
        pip install msal requests

    Environment variables:
        AZURE_CLIENT_ID  — App registration client ID
        AZURE_TENANT_ID  — Azure AD tenant ID (or 'common' for multi-tenant)
    """
    try:
        import msal
        import requests as req
    except ImportError:
        print("\n❌ Missing dependencies for OneDrive upload.")
        print("   Install with: pip install msal requests")
        return False

    client_id = os.environ.get("AZURE_CLIENT_ID", "")
    tenant_id = os.environ.get("AZURE_TENANT_ID", "common")

    if not client_id:
        print("\n❌ OneDrive upload requires AZURE_CLIENT_ID environment variable.")
        print("   Steps to set up:")
        print("   1. Go to https://portal.azure.com → Azure Active Directory → App registrations")
        print("   2. New registration → Name it → Set redirect URI to http://localhost")
        print("   3. API Permissions → Add → Microsoft Graph → Files.ReadWrite (delegated)")
        print("   4. Copy the Application (client) ID")
        print("   5. Set: export AZURE_CLIENT_ID='your-client-id'")
        print("   6. Set: export AZURE_TENANT_ID='your-tenant-id'  (optional, defaults to 'common')")
        return False

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["Files.ReadWrite"]

    # ── Authenticate with device code flow ──
    print("\n🔐 Authenticating with Microsoft...")
    app = msal.PublicClientApplication(client_id, authority=authority)

    # Try cached token first
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        # Device code flow (works without browser redirect issues)
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            print(f"❌ Failed to create device flow: {flow.get('error_description', 'Unknown error')}")
            return False

        print(f"\n   To authenticate, visit: {flow['verification_uri']}")
        print(f"   Enter code: {flow['user_code']}")
        print("   Waiting for authentication...")

        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print(f"❌ Authentication failed: {result.get('error_description', 'Unknown error')}")
        return False

    access_token = result["access_token"]
    print("   ✅ Authenticated successfully!")

    # ── Upload file ──
    file_name = os.path.basename(file_path)
    file_size = os.path.getsize(file_path)

    # Clean folder path
    folder_path = folder_path.strip("/")
    upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}/{file_name}:/content"

    print(f"\n☁️  Uploading to OneDrive: /{folder_path}/{file_name} ({file_size/1024:.1f} KB)...")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    }

    with open(file_path, "rb") as f:
        response = req.put(upload_url, headers=headers, data=f)

    if response.status_code in (200, 201):
        resp_data = response.json()
        web_url = resp_data.get("webUrl", "N/A")
        print(f"   ✅ Upload successful!")
        print(f"   📎 OneDrive URL: {web_url}")
        return True
    else:
        print(f"   ❌ Upload failed (HTTP {response.status_code})")
        print(f"   Response: {response.text[:500]}")
        return False


# ══════════════════════════════════════════════════════════════════════
# CLI ENTRY POINT
# ══════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Convert Splunk SPL Handbook (HTML) to PowerPoint (.pptx) and optionally upload to OneDrive.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python html_to_pptx.py                               # Convert only
  python html_to_pptx.py --upload                       # Convert + upload to OneDrive
  python html_to_pptx.py -i index.html -o handbook.pptx # Custom paths
  python html_to_pptx.py --upload --onedrive-folder "Presentations/SPL"
        """
    )
    parser.add_argument("-i", "--input", default="index.html",
                        help="Path to HTML file (default: index.html)")
    parser.add_argument("-o", "--output", default="Splunk_SPL_Handbook.pptx",
                        help="Output .pptx file path (default: Splunk_SPL_Handbook.pptx)")
    parser.add_argument("--upload", action="store_true",
                        help="Upload to OneDrive after conversion")
    parser.add_argument("--onedrive-folder", default="Documents",
                        help="OneDrive folder path (default: Documents)")

    args = parser.parse_args()

    # Validate input file
    if not os.path.isfile(args.input):
        print(f"❌ Input file not found: {args.input}")
        print(f"   Make sure '{args.input}' exists in the current directory.")
        sys.exit(1)

    print("=" * 60)
    print("  Splunk SPL Handbook → PowerPoint Converter")
    print("=" * 60)

    # Convert
    output_file = convert_html_to_pptx(args.input, args.output)

    # Upload
    if args.upload:
        print("\n" + "=" * 60)
        print("  OneDrive Upload")
        print("=" * 60)
        upload_to_onedrive(output_file, args.onedrive_folder)

    print("\n🎉 Done!")


if __name__ == "__main__":
    main()
