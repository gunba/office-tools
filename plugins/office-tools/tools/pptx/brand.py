"""Brand constants loader for the pptx tool.

Reads brand.toml (or brand.default.toml as a fallback) and exposes every
declared value as a module-level attribute. The pptx tool imports this module
and references constants like ``brand.ACCENT_PRIMARY`` or
``brand.TITLE_SIZE`` — the actual values come from the TOML.

Resolution order for the active TOML file:
  1. ``BRAND_TOML_PATH`` env var (Claude Code wires this through the plugin
     manifest's ``userConfig`` field — see ``plugin.json``).
  2. ``brand.toml`` next to this module, if present.
  3. ``brand.default.toml`` next to this module (always shipped).

TOML schema is documented in ``brand.default.toml``.
"""

from __future__ import annotations

import os
import sys
import tomllib
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.util import Pt, Inches

# ── Locate the active TOML file ──────────────────────────────────────────────

_HERE = Path(__file__).resolve().parent
_OVERRIDE_PATH_ENV = os.environ.get("BRAND_TOML_PATH", "").strip()
if _OVERRIDE_PATH_ENV:
    _BRAND_TOML = Path(_OVERRIDE_PATH_ENV).expanduser().resolve()
elif (_HERE / "brand.toml").is_file():
    _BRAND_TOML = _HERE / "brand.toml"
else:
    _BRAND_TOML = _HERE / "brand.default.toml"

if not _BRAND_TOML.is_file():
    raise FileNotFoundError(
        f"No brand TOML found. Looked for BRAND_TOML_PATH={_OVERRIDE_PATH_ENV!r}, "
        f"{_HERE / 'brand.toml'}, and {_HERE / 'brand.default.toml'}."
    )

with _BRAND_TOML.open("rb") as fh:
    _CFG = tomllib.load(fh)


def _hex_to_rgb(s: str) -> RGBColor:
    s = s.lstrip("#").strip()
    if len(s) != 6:
        raise ValueError(f"expected #RRGGBB hex color, got {s!r}")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


# ── Colors ───────────────────────────────────────────────────────────────────

_c = _CFG.get("colors", {})
ACCENT_PRIMARY = _hex_to_rgb(_c.get("accent_primary", "#1F3A66"))
ACCENT_SECONDARY = _hex_to_rgb(_c.get("accent_secondary", "#5C7A9A"))
DARK_NAVY = _hex_to_rgb(_c.get("dark_navy", "#0A1A33"))
BODY_TEXT = _hex_to_rgb(_c.get("body_text", "#3F4A55"))
LIGHT_GREY = _hex_to_rgb(_c.get("light_grey", "#CCCFCF"))
MEDIUM_GREY = _hex_to_rgb(_c.get("medium_grey", "#888B8D"))
MID_GREY_20 = _hex_to_rgb(_c.get("mid_grey_20", "#E7E7E8"))
PALE_BLUE = _hex_to_rgb(_c.get("pale_blue", "#E5F0F8"))
WHITE = _hex_to_rgb(_c.get("white", "#FFFFFF"))
BLACK = _hex_to_rgb(_c.get("black", "#000000"))

# ── Fonts ────────────────────────────────────────────────────────────────────

_f = _CFG.get("fonts", {})
FONT_FAMILY = _f.get("family", "Calibri")

# ── Sizes (Pt) ───────────────────────────────────────────────────────────────

_sz = _CFG.get("sizes", {})
TITLE_SIZE = Pt(_sz.get("title", 30))
SUBTITLE_SIZE = Pt(_sz.get("subtitle", 18))
SECTION_TITLE_SIZE = Pt(_sz.get("section_title", 44))
SECTION_NUM_SIZE = Pt(_sz.get("section_num", 96))
BODY_SIZE = Pt(_sz.get("body", 18))
SMALL_BODY_SIZE = Pt(_sz.get("small_body", 14))
PANEL_HEADER_SIZE = Pt(_sz.get("panel_header", 18))
PANEL_BODY_SIZE = Pt(_sz.get("panel_body", 14))
QUOTE_SIZE = Pt(_sz.get("quote", 28))
QUOTE_ATTR_SIZE = Pt(_sz.get("quote_attr", 14))
FOOTER_SIZE = Pt(_sz.get("footer", 9))
TABLE_HDR_SIZE = Pt(_sz.get("table_header", 11))
TABLE_BODY_SIZE = Pt(_sz.get("table_body", 10))

# ── Slide geometry (Inches) ──────────────────────────────────────────────────

_g = _CFG.get("slide", {})
SLIDE_W = Inches(_g.get("width", 13.333))
SLIDE_H = Inches(_g.get("height", 7.5))
MARGIN_L = Inches(_g.get("margin_l", 0.55))
MARGIN_R = Inches(_g.get("margin_r", 0.55))
MARGIN_T = Inches(_g.get("margin_t", 0.45))
MARGIN_B = Inches(_g.get("margin_b", 0.40))
TITLE_TOP = Inches(_g.get("title_top", 0.55))
TITLE_HEIGHT = Inches(_g.get("title_height", 0.65))
CONTENT_TOP = Inches(_g.get("content_top", 1.45))
CONTENT_H = Inches(_g.get("content_h", 5.5))
FOOTER_BAR_HEIGHT = Inches(_g.get("footer_bar_height", 0.05))
FOOTER_BAR_TOP = Inches(_g.get("footer_bar_top", 7.20))
FOOTER_TEXT_TOP = Inches(_g.get("footer_text_top", 7.28))
FOOTER_TEXT_H = Inches(_g.get("footer_text_h", 0.20))
ACCENT_RULE_TOP = Inches(_g.get("accent_rule_top", 1.25))
ACCENT_RULE_HEIGHT = Inches(_g.get("accent_rule_height", 0.04))
ACCENT_RULE_WIDTH = Inches(_g.get("accent_rule_width", 1.20))

# ── Logo ─────────────────────────────────────────────────────────────────────

_logo = _CFG.get("logo", {})
LOGO_PATH = _logo.get("path", "").strip()
LOGO_W = Inches(_logo.get("width_inches", 0.85))
LOGO_TOP = Inches(_logo.get("top_inches", 0.30))


def _summary() -> str:
    """Useful for debugging: which TOML was loaded."""
    return f"brand loaded from {_BRAND_TOML}"


if __name__ == "__main__":
    print(_summary(), file=sys.stderr)
