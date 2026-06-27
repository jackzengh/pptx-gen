"""Conservative text measurement so a blind model's text never overflows its box."""

import os
from functools import lru_cache

# EMU per inch (914400). Kept local so this module has no other harness deps.
EMU_PER_INCH = 914400

AVG_ADVANCE_EM = 0.55
POINTS_PER_INCH = 72
EMU_PER_POINT = EMU_PER_INCH / POINTS_PER_INCH

WRAP_FILL_FACTOR = 0.97

FONTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")

# basically a proxy for the real font files, we use Liberation fonts 
_BUNDLED_FACES = {
    "arial": "LiberationSans",
    "helvetica": "LiberationSans",
    "calibri": "LiberationSans",        
    "liberationsans": "LiberationSans",
    "timesnewroman": "LiberationSerif",
    "times": "LiberationSerif",
    "georgia": "LiberationSerif",
    "liberationserif": "LiberationSerif",
    "couriernew": "LiberationMono",
    "courier": "LiberationMono",
    "liberationmono": "LiberationMono",
}
_DEFAULT_FACE = "LiberationSans"


def _bundled_font_path(font_family: str, bold: bool = False, italic: bool = False) -> str | None:
    """Path to the bundled TTF (font file) for a requested family/style. this is how we approx text wrapping - using the ttf files! 

    Falls back to LiberationSans for any unknown family so the common Office/web
    fonts always resolve to a real file (no os.walk, no average-width guess).
    """
    family_key = "".join(ch for ch in (font_family or "").lower() if ch.isalnum())
    face = _BUNDLED_FACES.get(family_key, _DEFAULT_FACE)
    if bold and italic:
        suffix = "BoldItalic"
    elif bold:
        suffix = "Bold"
    elif italic:
        suffix = "Italic"
    else:
        suffix = "Regular"
    path = os.path.join(FONTS_DIR, f"{face}-{suffix}.ttf")
    return path if os.path.isfile(path) else None


@lru_cache(maxsize=256)
def _load_font(font_file: str, font_pt: float):
    from PIL import ImageFont

    return ImageFont.truetype(font_file, max(1, round(font_pt)))


def _load_resolved_font(
    font_family: str,
    font_file: str | None,
    bold: bool,
    italic: bool,
    font_pt: float,
):
    paths: list[str] = []
    # 1. Explicit theme-provided file wins (honored first).
    if font_file:
        path = os.path.expanduser(font_file)
        if os.path.isfile(path):
            paths.append(path)

    # 2. Bundled, metric-compatible twin of the requested family. 
    bundled = _bundled_font_path(font_family, bold, italic)
    if bundled and bundled not in paths:
        paths.append(bundled)

    # 3. Last resort: universal bundled fallback so we always have a real face.
    default_bundled = _bundled_font_path(_DEFAULT_FACE, bold, italic)
    if default_bundled and default_bundled not in paths:
        paths.append(default_bundled)

    for path in paths:
        try:
            return _load_font(path, font_pt)
        except Exception:
            continue
    return None


def measure_text_emu(
    text: str | list[str],
    box_width_emu: int,
    font_pt: float,
    leading_factor: float,
    *,
    space_after_pt: float = 0,
    font_family: str = "Arial",
    font_file: str | None = None,
    bold: bool = False,
    italic: bool = False,
    bold_items: list[bool] | None = None,
    italic_items: list[bool] | None = None,
    item_widths: list[int] | None = None,
) -> tuple[int, int]:
    # Return (width, height) in EMU for text content - this is how we approx whether text will wrap or not

    space_after_emu = round(space_after_pt * EMU_PER_POINT)
    is_single = isinstance(text, str)
    paragraphs = [text] if is_single else (text or [""])

    total_h = 0
    safe_box_w = max(1, int(box_width_emu))
    fallback_char_w_emu = max(1.0, AVG_ADVANCE_EM * font_pt * EMU_PER_POINT)
    fallback_line_h_emu = max(1, round(font_pt * leading_factor * EMU_PER_POINT))

    for i, para in enumerate(paragraphs):
        b = bold_items[i] if bold_items and i < len(bold_items) else bold
        it = italic_items[i] if italic_items and i < len(italic_items) else italic

        font = _load_resolved_font(font_family, font_file, b, it, font_pt)
        sample_text = para or " "

        if font is None:
            # No real font: estimate advance from a fixed average char width, but still
            # wrap word-by-word so a string just under the box width can't under-count.
            def measure_emu(s: str) -> int:
                return round(len(s) * fallback_char_w_emu)
            glyph_h_emu = fallback_line_h_emu
        else:
            def measure_emu(s: str) -> int:
                # getlength is the glyph ADVANCE width (cursor movement) — the right
                # metric for line wrapping, unlike getbbox's drawn-pixel extent.
                return round(font.getlength(s) * EMU_PER_POINT)
            ty_top, ty_bottom = font.getbbox("Ty")[1], font.getbbox("Ty")[3]
            glyph_h_emu = max(1, round((ty_bottom - ty_top) * EMU_PER_POINT))

        wrap_w = safe_box_w
        if item_widths and i < len(item_widths):
            wrap_w = max(1, int(item_widths[i]))
        wrapped_lines = _wrapped_line_count(sample_text, wrap_w, measure_emu)
        line_pitch_emu = max(glyph_h_emu, fallback_line_h_emu)
        total_h += wrapped_lines * line_pitch_emu

        if i < len(paragraphs) - 1:
            total_h += space_after_emu

    return safe_box_w, total_h


def _wrapped_line_count(text: str, box_w_emu: int, measure_emu) -> int:
    """Greedy word-wrap line count, mirroring how PowerPoint breaks lines.

    Accumulates words onto a line until the next word would exceed box width, then
    breaks — so end-of-line slack is honored (a string just under the box width that
    still can't fit its last word counts as 2 lines, not 1). A single word wider than
    the box is split into the ceil of its width over the box so an unbreakable token
    over-counts rather than under-counts. Always >= 1; biased to round up, never down.
    """
    # Conservatively shrink the usable width so a line that fills nearly the whole box
    # under the bundled twin still wraps under the user's (slightly wider) real font.
    usable = max(1, int(box_w_emu * WRAP_FILL_FACTOR))
    total = 0
    # Explicit newlines are hard breaks — each segment wraps independently, and an
    # empty segment (blank line) still occupies one line.
    for segment in text.split("\n"):
        words = segment.split()
        if not words:
            total += 1
            continue
        lines = 0
        current = ""
        for word in words:
            candidate = word if not current else f"{current} {word}"
            if measure_emu(candidate) <= usable:
                current = candidate
                continue
            # `word` doesn't fit after `current`. Close the current line (if any).
            if current:
                lines += 1
                current = ""
            # Place `word` on its own line; if it alone exceeds the box, it spans
            # multiple wrapped lines. (Use the full box for a lone over-long word so
            # a single long token isn't over-split.)
            word_w = measure_emu(word)
            if word_w <= usable:
                current = word
            else:
                lines += max(1, (word_w + box_w_emu - 1) // box_w_emu)
        if current:
            lines += 1
        total += max(1, lines)
    return max(1, total)
