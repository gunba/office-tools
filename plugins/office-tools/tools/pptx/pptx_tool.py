#!/usr/bin/env python3
"""
pptx_tool.py - Build, read, and edit PowerPoint (.pptx) decks.

Build  - generate a full deck from a JSON spec, styled with the active brand
         palette (read from brand.toml — see brand.py for the loader).
Read   - dump an existing .pptx to markdown for inspection.
Notes  - extract speaker notes only.

Usage:
    python tools/pptx/pptx_tool.py build  <spec.json> [--output deck.pptx]
    python tools/pptx/pptx_tool.py read   <file.pptx>
    python tools/pptx/pptx_tool.py notes  <file.pptx>

Spec format (JSON):
{
  "meta": {
    "title": "...", "subtitle": "...", "author": "...", "date": "...",
    "logo": "C:/path/to/logo.png",
    "footer_text": "Footer line | Subtitle | May 2026"
  },
  "slides": [
    {"layout": "title", "title": "...", "subtitle": "...", "notes": "..."},
    {"layout": "section", "number": "01", "title": "...", "notes": "..."},
    {"layout": "content", "title": "...", "bullets": ["...", "..."], "notes": "..."},
    {"layout": "two_column", "title": "...",
       "left_title": "...", "left_bullets": [...],
       "right_title": "...", "right_bullets": [...], "notes": "..."},
    {"layout": "three_panel", "title": "...",
       "panels": [
         {"title": "Chat", "bullets": [...]},
         {"title": "Agent", "bullets": [...]},
         {"layout": "Workflow", "bullets": [...]}
       ], "notes": "..."},
    {"layout": "image", "title": "...", "image": "path/to.png",
       "caption": "...", "notes": "..."},
    {"layout": "image_placeholder", "title": "...",
       "placeholder_text": "Architecture diagram",
       "image_prompt": "Full ChatGPT image prompt...", "notes": "..."},
    {"layout": "quote", "quote": "...", "attribution": "—...", "notes": "..."},
    {"layout": "table", "title": "...", "headers": [...], "rows": [[...], ...], "notes": "..."},
    {"layout": "big_text", "title": "...", "body": "Single paragraph", "notes": "..."}
  ]
}
"""

# === office-tools bootstrap ============================================
import os
import sys

_PLUGIN_ROOT = os.environ.get("CLAUDE_PLUGIN_ROOT") or os.environ.get("PLUGIN_ROOT")
if _PLUGIN_ROOT and _PLUGIN_ROOT not in sys.path:
    sys.path.insert(0, _PLUGIN_ROOT)
else:
    sys.path.insert(
        0,
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
    )

from bootstrap.ensure_python_deps import run as _ensure_python_deps  # noqa: E402

_ensure_python_deps()
# === end bootstrap =====================================================

import argparse
import json

from pptx import Presentation
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# Import brand constants from sibling module
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import brand as B  # noqa: E402


# ── Helpers ───────────────────────────────────────────────────────────────────

def _set_run(run, text, *, font=B.FONT_FAMILY, size=B.BODY_SIZE,
             color=B.BODY_TEXT, bold=False, italic=False):
    run.text = text
    run.font.name = font
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def _add_text(tf, text, *, size, color, bold=False, align=None,
              font=B.FONT_FAMILY, italic=False, replace_first=True):
    """Add a paragraph of text to a TextFrame.

    If replace_first=True and the textframe is empty (first paragraph), use it.
    """
    if replace_first and len(tf.paragraphs) == 1 and not tf.paragraphs[0].text:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    if align is not None:
        p.alignment = align
    run = p.add_run()
    _set_run(run, text, font=font, size=size, color=color, bold=bold, italic=italic)
    return p


def _shape_no_outline(shape):
    line = shape.line
    line.fill.background()


def _solid_fill(shape, color):
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_rect(slide, left, top, width, height, *, fill=None, no_line=True):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    if fill is not None:
        _solid_fill(rect, fill)
    if no_line:
        _shape_no_outline(rect)
    return rect


def _add_textbox(slide, left, top, width, height):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)
    return tb, tf


def _set_notes(slide, text):
    if not text:
        return
    notes_tf = slide.notes_slide.notes_text_frame
    notes_tf.text = text


def _add_chrome(slide, *, page_num, total, footer_text, logo_path):
    """Add header logo + footer bar/page number to a slide."""
    # Footer accent bar
    _add_rect(slide,
              left=Emu(0), top=B.FOOTER_BAR_TOP,
              width=B.SLIDE_W, height=B.FOOTER_BAR_HEIGHT,
              fill=B.ACCENT_PRIMARY)

    # Footer text (left) + page number (right)
    _, lf = _add_textbox(slide,
                         left=B.MARGIN_L, top=B.FOOTER_TEXT_TOP,
                         width=Inches(8.0), height=B.FOOTER_TEXT_H)
    _add_text(lf, footer_text, size=B.FOOTER_SIZE, color=B.MEDIUM_GREY,
              align=PP_ALIGN.LEFT)

    _, rf = _add_textbox(slide,
                         left=Inches(10.5), top=B.FOOTER_TEXT_TOP,
                         width=Inches(2.30), height=B.FOOTER_TEXT_H)
    _add_text(rf, f"{page_num} / {total}", size=B.FOOTER_SIZE,
              color=B.MEDIUM_GREY, align=PP_ALIGN.RIGHT)

    # Logo top-right (skip on title and section dividers — handled by caller)
    if logo_path and os.path.exists(logo_path):
        try:
            slide.shapes.add_picture(
                logo_path,
                left=B.SLIDE_W - B.MARGIN_R - B.LOGO_W,
                top=B.LOGO_TOP,
                width=B.LOGO_W,
            )
        except Exception:
            pass


def _add_slide_title(slide, title_text):
    """Title text + accent rule beneath."""
    title_w = B.SLIDE_W - B.MARGIN_L - B.MARGIN_R - B.LOGO_W - Inches(0.2)
    _, tf = _add_textbox(slide,
                         left=B.MARGIN_L, top=B.TITLE_TOP,
                         width=title_w, height=B.TITLE_HEIGHT)
    _add_text(tf, title_text, size=B.TITLE_SIZE, color=B.DARK_NAVY, bold=True)

    # Accent rule under title
    _add_rect(slide,
              left=B.MARGIN_L, top=B.ACCENT_RULE_TOP,
              width=B.ACCENT_RULE_WIDTH, height=B.ACCENT_RULE_HEIGHT,
              fill=B.ACCENT_PRIMARY)


# ── Slide builders ────────────────────────────────────────────────────────────

def build_title_slide(prs, spec, meta):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    # Full-bleed pale band on the left as a brand accent
    _add_rect(slide, left=Emu(0), top=Emu(0),
              width=Inches(0.45), height=B.SLIDE_H,
              fill=B.ACCENT_PRIMARY)

    # Logo, larger, centered-ish top
    logo_path = meta.get("logo")
    if logo_path and os.path.exists(logo_path):
        try:
            slide.shapes.add_picture(
                logo_path,
                left=Inches(1.0), top=Inches(0.7),
                width=Inches(1.6),
            )
        except Exception:
            pass

    # Title
    _, tf = _add_textbox(slide,
                         left=Inches(1.0), top=Inches(2.6),
                         width=Inches(11.5), height=Inches(2.0))
    _add_text(tf, spec["title"], size=Pt(48), color=B.DARK_NAVY, bold=True)

    # Subtitle
    if spec.get("subtitle"):
        _, sf = _add_textbox(slide,
                             left=Inches(1.0), top=Inches(4.4),
                             width=Inches(11.5), height=Inches(0.8))
        _add_text(sf, spec["subtitle"], size=B.SUBTITLE_SIZE, color=B.ACCENT_PRIMARY)

    # Author + date footer
    foot_parts = [meta.get("author"), meta.get("date")]
    foot = " | ".join(p for p in foot_parts if p)
    if foot:
        _, ff = _add_textbox(slide,
                             left=Inches(1.0), top=Inches(6.6),
                             width=Inches(11.5), height=Inches(0.4))
        _add_text(ff, foot, size=Pt(12), color=B.MEDIUM_GREY)

    _set_notes(slide, spec.get("notes", ""))
    return slide


def build_section_slide(prs, spec, meta):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    # Dark navy background band on left
    _add_rect(slide, left=Emu(0), top=Emu(0),
              width=Inches(4.5), height=B.SLIDE_H,
              fill=B.DARK_NAVY)

    # Big section number
    if spec.get("number"):
        _, nf = _add_textbox(slide,
                             left=Inches(0.6), top=Inches(2.0),
                             width=Inches(3.5), height=Inches(2.0))
        _add_text(nf, spec["number"], size=B.SECTION_NUM_SIZE,
                  color=B.ACCENT_PRIMARY, bold=True)

    # Section title (right side)
    _, tf = _add_textbox(slide,
                         left=Inches(5.0), top=Inches(3.0),
                         width=Inches(7.8), height=Inches(2.0))
    _add_text(tf, spec["title"], size=B.SECTION_TITLE_SIZE,
              color=B.DARK_NAVY, bold=True)

    # Subtitle
    if spec.get("subtitle"):
        _, sf = _add_textbox(slide,
                             left=Inches(5.0), top=Inches(4.5),
                             width=Inches(7.8), height=Inches(1.0))
        _add_text(sf, spec["subtitle"], size=Pt(20), color=B.BODY_TEXT)

    _set_notes(slide, spec.get("notes", ""))
    return slide


def build_content_slide(prs, spec, meta, page_num, total):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    _add_chrome(slide, page_num=page_num, total=total,
                footer_text=meta.get("footer_text", ""),
                logo_path=meta.get("logo"))
    _add_slide_title(slide, spec["title"])

    # Body bullets
    bullets = spec.get("bullets", [])
    body_w = B.SLIDE_W - B.MARGIN_L - B.MARGIN_R
    body_h = B.CONTENT_H

    _, bf = _add_textbox(slide,
                         left=B.MARGIN_L, top=B.CONTENT_TOP,
                         width=body_w, height=body_h)

    body_size_raw = spec.get("body_size")
    if body_size_raw is None:
        body_size = B.BODY_SIZE
    else:
        body_size = Pt(int(body_size_raw))

    for i, item in enumerate(bullets):
        # bullet may be a string or {text, sub: [...]}
        if isinstance(item, dict):
            text = item.get("text", "")
            subs = item.get("sub", [])
        else:
            text = str(item)
            subs = []

        p = _add_text(bf, f"• {text}", size=body_size, color=B.DARK_NAVY,
                      bold=False, replace_first=(i == 0))
        p.space_after = Pt(8)
        p.line_spacing = 1.15

        for sub in subs:
            sp = _add_text(bf, f"   – {sub}", size=Pt(14), color=B.BODY_TEXT,
                           replace_first=False)
            sp.space_after = Pt(4)
            sp.line_spacing = 1.15

    _set_notes(slide, spec.get("notes", ""))
    return slide


def build_two_column_slide(prs, spec, meta, page_num, total):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    _add_chrome(slide, page_num=page_num, total=total,
                footer_text=meta.get("footer_text", ""),
                logo_path=meta.get("logo"))
    _add_slide_title(slide, spec["title"])

    col_w = Emu(int((B.SLIDE_W - B.MARGIN_L - B.MARGIN_R - Inches(0.5)) / 2))

    # Left column
    _, lt = _add_textbox(slide,
                         left=B.MARGIN_L, top=B.CONTENT_TOP,
                         width=col_w, height=Inches(0.5))
    _add_text(lt, spec.get("left_title", ""), size=Pt(20),
              color=B.ACCENT_PRIMARY, bold=True)

    _, lb = _add_textbox(slide,
                         left=B.MARGIN_L, top=B.CONTENT_TOP + Inches(0.55),
                         width=col_w, height=B.CONTENT_H - Inches(0.55))
    for i, item in enumerate(spec.get("left_bullets", [])):
        p = _add_text(lb, f"• {item}", size=B.SMALL_BODY_SIZE,
                      color=B.DARK_NAVY, replace_first=(i == 0))
        p.space_after = Pt(6)
        p.line_spacing = 1.15

    # Right column
    right_left = B.MARGIN_L + col_w + Inches(0.5)
    _, rt = _add_textbox(slide,
                         left=right_left, top=B.CONTENT_TOP,
                         width=col_w, height=Inches(0.5))
    _add_text(rt, spec.get("right_title", ""), size=Pt(20),
              color=B.ACCENT_SECONDARY, bold=True)

    _, rb = _add_textbox(slide,
                         left=right_left, top=B.CONTENT_TOP + Inches(0.55),
                         width=col_w, height=B.CONTENT_H - Inches(0.55))
    for i, item in enumerate(spec.get("right_bullets", [])):
        p = _add_text(rb, f"• {item}", size=B.SMALL_BODY_SIZE,
                      color=B.DARK_NAVY, replace_first=(i == 0))
        p.space_after = Pt(6)
        p.line_spacing = 1.15

    # Vertical divider
    _add_rect(slide,
              left=B.MARGIN_L + col_w + Inches(0.225), top=B.CONTENT_TOP,
              width=Inches(0.05), height=B.CONTENT_H,
              fill=B.LIGHT_GREY)

    _set_notes(slide, spec.get("notes", ""))
    return slide


def build_three_panel_slide(prs, spec, meta, page_num, total):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    _add_chrome(slide, page_num=page_num, total=total,
                footer_text=meta.get("footer_text", ""),
                logo_path=meta.get("logo"))
    _add_slide_title(slide, spec["title"])

    panels = spec.get("panels", [])
    if not panels:
        return slide

    n = len(panels)
    gap = Inches(0.25)
    avail_w = B.SLIDE_W - B.MARGIN_L - B.MARGIN_R - gap * (n - 1)
    panel_w = Emu(int(avail_w / n))
    panel_h = B.CONTENT_H
    accent_colors = [B.ACCENT_PRIMARY, B.ACCENT_SECONDARY, B.DARK_NAVY, B.MEDIUM_GREY]

    for i, panel in enumerate(panels):
        left = B.MARGIN_L + (panel_w + gap) * i
        # Panel background (very light)
        _add_rect(slide, left=left, top=B.CONTENT_TOP,
                  width=panel_w, height=panel_h,
                  fill=B.MID_GREY_20)
        # Top accent bar
        _add_rect(slide, left=left, top=B.CONTENT_TOP,
                  width=panel_w, height=Inches(0.08),
                  fill=accent_colors[i % len(accent_colors)])
        # Header
        _, hf = _add_textbox(slide,
                             left=left + Inches(0.2),
                             top=B.CONTENT_TOP + Inches(0.2),
                             width=panel_w - Inches(0.4),
                             height=Inches(0.6))
        _add_text(hf, panel.get("title", ""), size=B.PANEL_HEADER_SIZE,
                  color=B.DARK_NAVY, bold=True)

        # Body bullets
        _, bf = _add_textbox(slide,
                             left=left + Inches(0.2),
                             top=B.CONTENT_TOP + Inches(0.85),
                             width=panel_w - Inches(0.4),
                             height=panel_h - Inches(1.0))
        for j, b in enumerate(panel.get("bullets", [])):
            p = _add_text(bf, f"• {b}", size=B.PANEL_BODY_SIZE,
                          color=B.DARK_NAVY, replace_first=(j == 0))
            p.space_after = Pt(4)
            p.line_spacing = 1.15

    _set_notes(slide, spec.get("notes", ""))
    return slide


def build_image_slide(prs, spec, meta, page_num, total):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    _add_chrome(slide, page_num=page_num, total=total,
                footer_text=meta.get("footer_text", ""),
                logo_path=meta.get("logo"))
    _add_slide_title(slide, spec["title"])

    img_path = spec.get("image")
    if img_path and os.path.exists(img_path):
        # Center image in the content area
        img_max_w = B.SLIDE_W - B.MARGIN_L - B.MARGIN_R
        img_max_h = B.CONTENT_H - Inches(0.4)  # leave room for caption
        try:
            slide.shapes.add_picture(
                img_path,
                left=B.MARGIN_L + Inches(0.5),
                top=B.CONTENT_TOP,
                width=img_max_w - Inches(1.0),
            )
        except Exception:
            pass

    if spec.get("caption"):
        _, cf = _add_textbox(slide,
                             left=B.MARGIN_L,
                             top=B.CONTENT_TOP + B.CONTENT_H - Inches(0.4),
                             width=B.SLIDE_W - B.MARGIN_L - B.MARGIN_R,
                             height=Inches(0.4))
        _add_text(cf, spec["caption"], size=Pt(12),
                  color=B.MEDIUM_GREY, italic=True, align=PP_ALIGN.CENTER)

    _set_notes(slide, spec.get("notes", ""))
    return slide


def build_image_placeholder_slide(prs, spec, meta, page_num, total):
    """Same chrome as image slide, but shows a styled placeholder rectangle.

    The full image-generation prompt is appended to the speaker notes so the
    presenter can paste it straight into ChatGPT/Claude Image and drop the
    result into the slide later.
    """
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    _add_chrome(slide, page_num=page_num, total=total,
                footer_text=meta.get("footer_text", ""),
                logo_path=meta.get("logo"))
    _add_slide_title(slide, spec["title"])

    # Placeholder rectangle
    box_left = B.MARGIN_L + Inches(0.5)
    box_top  = B.CONTENT_TOP
    box_w    = B.SLIDE_W - B.MARGIN_L - B.MARGIN_R - Inches(1.0)
    box_h    = B.CONTENT_H - Inches(0.5)

    rect = _add_rect(slide, left=box_left, top=box_top,
                     width=box_w, height=box_h,
                     fill=B.PALE_BLUE)
    # Dashed border
    line = rect.line
    line.color.rgb = B.ACCENT_PRIMARY
    line.width = Pt(1.5)
    # Dash style requires direct XML
    spPr = rect.line._get_or_add_ln()
    prstDash = etree.SubElement(spPr, qn('a:prstDash'))
    prstDash.set('val', 'dash')

    placeholder_text = spec.get("placeholder_text", "Image to be inserted")
    _, pf = _add_textbox(slide,
                         left=box_left, top=box_top + box_h / 2 - Inches(0.4),
                         width=box_w, height=Inches(0.8))
    _add_text(pf, f"[ {placeholder_text} ]", size=Pt(20),
              color=B.ACCENT_PRIMARY, bold=True, align=PP_ALIGN.CENTER)
    _add_text(pf, "Image generation prompt is in the speaker notes below.",
              size=Pt(12), color=B.MEDIUM_GREY, italic=True,
              align=PP_ALIGN.CENTER, replace_first=False)

    # Notes: existing notes + image prompt
    notes_parts = []
    if spec.get("notes"):
        notes_parts.append(spec["notes"].strip())
    if spec.get("image_prompt"):
        notes_parts.append("\n--- IMAGE PROMPT ---\n")
        notes_parts.append(spec["image_prompt"].strip())
    if notes_parts:
        _set_notes(slide, "\n".join(notes_parts))
    return slide


def build_quote_slide(prs, spec, meta, page_num, total):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    _add_chrome(slide, page_num=page_num, total=total,
                footer_text=meta.get("footer_text", ""),
                logo_path=meta.get("logo"))

    # Big quote, vertically centered
    _, qf = _add_textbox(slide,
                         left=Inches(1.5), top=Inches(2.0),
                         width=Inches(10.3), height=Inches(3.5))
    _add_text(qf, f"“{spec['quote']}”", size=B.QUOTE_SIZE,
              color=B.DARK_NAVY, italic=True)

    if spec.get("attribution"):
        _, af = _add_textbox(slide,
                             left=Inches(1.5), top=Inches(5.7),
                             width=Inches(10.3), height=Inches(0.5))
        _add_text(af, f"— {spec['attribution']}", size=B.QUOTE_ATTR_SIZE,
                  color=B.ACCENT_PRIMARY)

    _set_notes(slide, spec.get("notes", ""))
    return slide


def build_table_slide(prs, spec, meta, page_num, total):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    _add_chrome(slide, page_num=page_num, total=total,
                footer_text=meta.get("footer_text", ""),
                logo_path=meta.get("logo"))
    _add_slide_title(slide, spec["title"])

    headers = spec.get("headers", [])
    rows = spec.get("rows", [])
    if not headers and not rows:
        return slide

    n_rows = 1 + len(rows)
    n_cols = len(headers) if headers else max(len(r) for r in rows)

    table_w = B.SLIDE_W - B.MARGIN_L - B.MARGIN_R
    table_h_spec = spec.get("table_height")
    if table_h_spec == "full":
        table_h = B.CONTENT_H
    elif table_h_spec is not None:
        table_h = Inches(float(table_h_spec))
    else:
        table_h = min(B.CONTENT_H, Inches(0.5) + Inches(0.4) * len(rows))

    table_body_size = B.TABLE_BODY_SIZE
    if spec.get("table_body_size") is not None:
        table_body_size = Pt(int(spec["table_body_size"]))

    table_shape = slide.shapes.add_table(
        n_rows, n_cols,
        B.MARGIN_L, B.CONTENT_TOP,
        table_w, table_h,
    )
    tbl = table_shape.table

    # Header row
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = B.DARK_NAVY
        tf = cell.text_frame
        tf.text = ""
        _add_text(tf, h, size=B.TABLE_HDR_SIZE, color=B.WHITE, bold=True)

    # Body
    for i, row in enumerate(rows, start=1):
        alt = (i % 2 == 0)
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            if alt:
                cell.fill.solid()
                cell.fill.fore_color.rgb = B.MID_GREY_20
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = B.WHITE
            tf = cell.text_frame
            tf.text = ""
            _add_text(tf, str(val), size=table_body_size, color=B.DARK_NAVY)

    _set_notes(slide, spec.get("notes", ""))
    return slide


def build_big_text_slide(prs, spec, meta, page_num, total):
    """Title + a single large statement. For high-impact takeaway slides."""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    _add_chrome(slide, page_num=page_num, total=total,
                footer_text=meta.get("footer_text", ""),
                logo_path=meta.get("logo"))
    if spec.get("title"):
        _add_slide_title(slide, spec["title"])

    body_top = B.CONTENT_TOP if spec.get("title") else Inches(2.5)
    body_h = B.CONTENT_H if spec.get("title") else Inches(3.0)

    _, bf = _add_textbox(slide,
                         left=Inches(1.5), top=body_top,
                         width=Inches(10.3), height=body_h)
    _add_text(bf, spec.get("body", ""), size=Pt(28),
              color=B.DARK_NAVY, bold=False)

    _set_notes(slide, spec.get("notes", ""))
    return slide


# ── Dispatcher ────────────────────────────────────────────────────────────────

LAYOUT_BUILDERS = {
    "title":             build_title_slide,
    "section":           build_section_slide,
    "content":           build_content_slide,
    "two_column":        build_two_column_slide,
    "three_panel":       build_three_panel_slide,
    "image":             build_image_slide,
    "image_placeholder": build_image_placeholder_slide,
    "quote":             build_quote_slide,
    "table":             build_table_slide,
    "big_text":          build_big_text_slide,
}


# ── Commands ──────────────────────────────────────────────────────────────────

def cmd_build(args):
    with open(args.spec, "r", encoding="utf-8") as fh:
        spec = json.load(fh)

    meta = spec.get("meta", {})
    slides = spec.get("slides", [])

    prs = Presentation()
    prs.slide_width = B.SLIDE_W
    prs.slide_height = B.SLIDE_H

    total = len(slides)
    for i, sl in enumerate(slides, start=1):
        layout = sl.get("layout", "content")
        builder = LAYOUT_BUILDERS.get(layout)
        if not builder:
            print(f"  [skip] slide {i}: unknown layout '{layout}'", file=sys.stderr)
            continue
        if layout in ("title", "section"):
            builder(prs, sl, meta)
        else:
            builder(prs, sl, meta, i, total)
        print(f"  [ok]   slide {i}/{total}: {layout} - {sl.get('title', '')[:60]}")

    output = args.output or os.path.splitext(args.spec)[0] + ".pptx"
    prs.save(output)
    print(f"\nSaved: {output}  ({total} slides)")


def cmd_read(args):
    prs = Presentation(args.file)
    out = []
    for i, slide in enumerate(prs.slides, start=1):
        out.append(f"\n## Slide {i}")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text
                    if text.strip():
                        out.append(text)
            elif shape.shape_type == 19:  # TABLE
                tbl = shape.table
                rows_md = []
                for ri, row in enumerate(tbl.rows):
                    cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                    rows_md.append("| " + " | ".join(cells) + " |")
                    if ri == 0:
                        rows_md.append("| " + " | ".join(["---"] * len(cells)) + " |")
                out.append("")
                out.extend(rows_md)
                out.append("")
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                out.append("\n_Speaker notes:_")
                out.append(notes)
    print("\n".join(out).strip())


def cmd_notes(args):
    prs = Presentation(args.file)
    for i, slide in enumerate(prs.slides, start=1):
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                print(f"\n## Slide {i}")
                print(notes)


def main():
    parser = argparse.ArgumentParser(description="Build, read, or extract notes from .pptx")
    sub = parser.add_subparsers(dest="command")

    p_build = sub.add_parser("build", help="Build a deck from a JSON spec")
    p_build.add_argument("spec")
    p_build.add_argument("--output", "-o", help="Output .pptx path")

    p_read = sub.add_parser("read", help="Read deck contents as text")
    p_read.add_argument("file")

    p_notes = sub.add_parser("notes", help="Extract speaker notes only")
    p_notes.add_argument("file")

    args = parser.parse_args()

    commands = {"build": cmd_build, "read": cmd_read, "notes": cmd_notes}
    if args.command in commands:
        commands[args.command](args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
