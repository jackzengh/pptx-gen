# python-pptx Build Reference

How to build a `.pptx` from scratch with **python-pptx** (`pip install python-pptx`, import name `pptx`). This is the python equivalent of the pptxgenjs guide, mapped to the verified python-pptx API.

Two rules run through everything here (see [SKILL.md](SKILL.md) and [naming.md](naming.md)):

1. **Every text frame gets `tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`.** This makes the box grow to fit its text when PowerPoint/LibreOffice renders it, so content isn't clipped. Note: this is an authoring best practice — it does **not** make `check_pptx` see in-box overflow (the validator reads stored geometry, which `auto_size` doesn't change). Catch overflow with a visual render, not `check_pptx`.
2. **Every shape gets a meaningful `shape.name`.** `check_pptx` echoes names in every problem it reports, so good names make failures self-explaining.

---

## Setup & Basic Structure

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

prs = Presentation()

# 16:9 (a fresh Presentation() defaults to 4:3 / 10 x 7.5in)
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)

# Blank layout is index 6 in the bundled default template
# (0=Title, 1=Title+Content, 5=Title Only, 6=Blank).
slide = prs.slides.add_slide(prs.slide_layouts[6])

prs.save("deck.pptx")
```

- The origin is the **top-left** of the slide; `left`/`top` increase right/down.
- Positions and sizes are `Length` objects, never bare numbers: `Inches(1)`, `Pt(18)`, `Emu(914400)`. `1 inch = 914400 EMU`, `1 pt = 12700 EMU`.

### Layout dimensions

| Aspect | Width × Height |
|--------|----------------|
| 16:9   | `Inches(13.333)` × `Inches(7.5)` |
| 16:10  | `Inches(12)` × `Inches(7.5)` |
| 4:3    | `Inches(10)` × `Inches(7.5)` (default) |

`check_pptx` reads the real `prs.slide_width`/`slide_height`, so its margin checks adapt to whatever size you set.

---

## Text & Formatting

```python
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1.2))
box.name = "title"                                   # always name it
tf = box.text_frame

tf.word_wrap = True
tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT       # rule 1: avoid clipping at render time
tf.vertical_anchor = MSO_ANCHOR.MIDDLE               # TOP / MIDDLE / BOTTOM

# Paragraph 0 already exists
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT                          # LEFT / CENTER / RIGHT / JUSTIFY
p.text = "Quarterly Results"

run = p.runs[0]
run.font.name = "Calibri"
run.font.size = Pt(40)
run.font.bold = True
run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

# More lines: add paragraphs, then runs for per-span formatting
p2 = tf.add_paragraph()
r2 = p2.add_run()
r2.text = "Fiscal year 2025"
r2.font.size = Pt(16)
```

- `from pptx.enum.text import MSO_AUTO_SIZE` — members: `NONE`, `SHAPE_TO_FIT_TEXT`, `TEXT_TO_FIT_SHAPE`.
- `tf.text = "..."` is a shortcut that replaces all content with a single paragraph/run.
- **Internal margins**: text frames have padding by default. To align text precisely with shapes/lines at the same x, zero it out:

```python
tf.margin_left = Pt(0)
tf.margin_right = Pt(0)
tf.margin_top = Pt(0)
tf.margin_bottom = Pt(0)
```

> Colors are `RGBColor(0xRR, 0xGG, 0xBB)` — **never** a `"#RRGGBB"` string.

---

## Lists & Bullets

`paragraph.level` (0–8) only sets the **indent / outline level** — it does not draw a bullet glyph.

Bullets are driven by **`p.level`** (the 0-based nesting level, 0–8) plus a real bullet glyph (`<a:buChar>`) and a **hanging indent** so wrapped lines align under the text, not under the glyph. python-pptx has no `bullet=True` toggle, so use the helper below.

- **Body placeholders** inherited from a layout usually already carry bullet formatting; plain textboxes do not — that's the case the helper handles.
- Never type a literal `"• "` into the run **and** apply a bullet — you get a double bullet.

See **Reusable helpers** below for `_apply_bullet` (real glyphs + hanging indent keyed off `p.level`) and `paragraph(..., bullet=True, level=n)`.

---

## Reusable helpers (recommended)

Build every textbox, bulleted paragraph, rectangle, and connector through these helpers so geometry, naming, and bullet indentation stay consistent across the deck. All coordinates are **EMU** (`Emu(...)`); `Inches`/`Pt` convert in.

```python
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

_EMU_PER_POINT = 914400 / 72

# Bullet glyph per nesting level (0-based); clamp beyond the last.
_BULLET_GLYPHS = ["•", "–", "◦"]
# Hanging-indent geometry, in em units of the font size:
#   _HANG_EM — distance from the glyph to the text (the hang)
#   _STEP_EM — extra left margin added per nesting level
_HANG_EM = 1.1
_STEP_EM = 1.4


def bullet_indent_emu(level: int, size_pt: float) -> tuple[int, int]:
    """(marL, indent) in EMU for a bullet at this level/size.

    marL is the paragraph's left margin (where the text column starts); indent is the
    first-line indent (negative -> the glyph hangs left of the text). One source of truth
    so the draw side and any measure side agree on the wrap width.
    """
    em = size_pt * _EMU_PER_POINT
    marL = round((_HANG_EM + _STEP_EM * max(0, level)) * em)
    indent = round(-_HANG_EM * em)   # negative = hanging indent
    return marL, indent


def _apply_bullet(p, level: int, size_pt: float, font: str) -> None:
    """Give paragraph `p` a real bullet glyph with a hanging indent (idempotent)."""
    marL, indent = bullet_indent_emu(level, size_pt)
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(marL))
    pPr.set("indent", str(indent))
    # Drop any existing bullet definition so re-application is clean.
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buFont"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    glyph = _BULLET_GLYPHS[min(max(0, level), len(_BULLET_GLYPHS) - 1)]
    pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": font}))
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": glyph}))


def style_run(run, size_pt, color: RGBColor, *, bold=False, italic=False, font="Arial"):
    run.font.size = Pt(size_pt)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font


def textbox(slide, left: int, top: int, width: int, height: int, *, name: str):
    """Add a zero-padding textbox at EMU bounds. Returns (shape, text_frame)."""
    tb = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT   # required by the skill
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    return tb, tf


def paragraph(tf, text, size_pt, color, *, bold=False, italic=False, level=0,
              align=PP_ALIGN.LEFT, space_after_pt=6, space_before_pt=0,
              first=False, font="Arial", bullet=False):
    """Add (or set the first) paragraph with consistent styling. `bullet` is opt-in."""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.level = level
    p.alignment = align
    p.space_after = Pt(space_after_pt)
    p.space_before = Pt(space_before_pt)
    if bullet:
        _apply_bullet(p, level, size_pt, font)
    run = p.add_run()
    run.text = text
    style_run(run, size_pt, color, bold=bold, italic=italic, font=font)
    return p


def rectangle(slide, left: int, top: int, width: int, height: int, *, name: str,
              fill: RGBColor | None = None, line: RGBColor | None = None, line_w_pt=0.75):
    """Rectangle at EMU bounds. fill/line None means transparent."""
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(left), Emu(top), Emu(width), Emu(height))
    sp.name = name
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w_pt)
    return sp


def connector(slide, left, top, width, color, *, name: str, weight_pt=1.5):
    """Horizontal rule (connector) of the given EMU width. (MSO_CONNECTOR.STRAIGHT == 2)"""
    line = slide.shapes.add_connector(2, Emu(left), Emu(top), Emu(left + width), Emu(top))
    line.name = name
    line.line.color.rgb = color
    line.line.width = Pt(weight_pt)
    return line
```

Usage — a bulleted list with sub-points (note `level=` drives both the glyph and the indent):

```python
_, tf = textbox(slide, Emu(MARGIN), Emu(top), Emu(col_w), Emu(col_h), name="body_highlights")
paragraph(tf, "RCS grew 50% year over year.", 15, INK, first=True, bullet=True, level=0)
paragraph(tf, "Became the majority of revenue.", 13, MUTED, bullet=True, level=1)
paragraph(tf, "Up from 45% a year earlier.", 13, MUTED, bullet=True, level=1)
```

> **auto_size note:** `textbox()` sets `tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT` per the skill. This makes the box grow to fit its text when rendered, so content isn't clipped. It does **not** change the stored geometry `check_pptx` reads — so still pass a generous `height` and verify overflow with a visual render (see [SKILL.md](SKILL.md)).

---

## Shapes

```python
from pptx.enum.shapes import MSO_SHAPE

card = slide.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(2), Inches(3.5), Inches(2)
)
card.name = "card_revenue"                    # always name it

card.fill.solid()
card.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)

card.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
card.line.width = Pt(1)

# No fill / no border
# card.fill.background()        # transparent interior
# card.line.fill.background()   # no border
```

Common `MSO_SHAPE` members: `RECTANGLE`, `ROUNDED_RECTANGLE`, `OVAL`, `DIAMOND`, `CHEVRON`, `RIGHT_ARROW`.

**Lines are not autoshapes** — use a connector:

```python
from pptx.enum.shapes import MSO_CONNECTOR
line = slide.shapes.add_connector(
    MSO_CONNECTOR.STRAIGHT, Inches(1), Inches(4.5), Inches(5), Inches(4.5)
)
line.name = "divider_s1"
line.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
line.line.width = Pt(1.5)
```

### Transparency (not native)

python-pptx has no transparency/alpha option. `fill.solid()` is opaque. If you truly need a semi-transparent fill, inject an `<a:alpha>` element (val is 0–100000, so `50000` = 50%):

```python
from pptx.oxml.ns import qn
card.fill.solid()
srgb = card.fill.fore_color._xFill.find(qn('a:srgbClr'))
srgb.append(srgb.makeelement(qn('a:alpha'), {'val': '50000'}))
```

Prefer choosing a lighter solid color instead — it avoids the XML entirely.

---

## Backgrounds: panel shape vs. true slide background

This distinction matters for `check_pptx`, which only ever walks `slide.shapes`.

**Full-bleed slide background** — lives in `<p:bg>`, is **NOT a shape**, and `check_pptx` never sees it (so a full-bleed background can't trip overlap/margin checks):

```python
bg = slide.background
bg.fill.solid()
bg.fill.fore_color.rgb = RGBColor(0x0B, 0x0B, 0x1A)
```

> Reading `slide.background.fill` is destructive — it discards any inherited background and replaces it. That's fine when you're setting it.

**A background band/panel that should sit in the layout** (e.g. a colored sidebar or header strip you want aligned and non-overlapping) — make it a **named shape**:

```python
band = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.2)
)
band.name = "background_header"        # has "background" in the name (see naming.md)
band.fill.solid()
band.fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)
band.line.fill.background()
```

A band like this has no text frame, so `check_pptx` classifies it as `kind="other"`: it is **skipped** by alignment / left-margin / gutter checks, but **still counted** by overlap and outer-margin checks. Keep panels off the edge band and behind (added before) the content that overlays them — note that two filled boxes overlapping *will* be flagged, so for content sitting on a panel, prefer the true `<p:bg>` background or place content beside the panel rather than on top of it.

---

## Images

```python
pic = slide.shapes.add_picture(
    "chart.png", Inches(7), Inches(1.5), width=Inches(5), height=Inches(3)
)
pic.name = "picture_trend"
```

Omit `width`/`height` for native size, or pass one to scale proportionally.

---

## Tables

```python
gf = slide.shapes.add_table(
    rows=3, cols=2, left=Inches(1), top=Inches(2), width=Inches(6), height=Inches(2)
)
gf.name = "table_pricing"
table = gf.table

table.cell(0, 0).text = "Plan"
table.cell(0, 1).text = "Price"
table.cell(0, 0).text_frame.paragraphs[0].font.bold = True

table.columns[0].width = Inches(4)
table.rows[0].height = Inches(0.6)

# Merge a rectangular region (origin cell .merge(opposite corner))
table.cell(0, 0).merge(table.cell(0, 1))
```

> Give every table a caption textbox within 0.1in **above** it — see Charts below; the same `element-boxing` rule applies to tables.

---

## Charts

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

data = CategoryChartData()
data.categories = ["Q1", "Q2", "Q3", "Q4"]
data.add_series("Revenue", (45, 55, 62, 71))

gf = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(2.1), Inches(6), Inches(3.5), data,
)
gf.name = "chart_revenue"
chart = gf.chart

chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM

series = chart.plots[0].series[0]
series.format.fill.solid()
series.format.fill.fore_color.rgb = RGBColor(0x4C, 0xAF, 0x50)
```

Common `XL_CHART_TYPE`: `COLUMN_CLUSTERED`, `BAR_CLUSTERED`, `LINE`, `LINE_MARKERS`, `PIE`, `DOUGHNUT`, `XY_SCATTER`, `AREA`. Use `CategoryChartData` for column/bar/line/pie.

> **`check_pptx` requires a caption above every chart and table.** Add a text box whose **bottom** sits within 0.1in of the chart's **top**, horizontally overlapping it:

```python
cap = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(6), Inches(0.5))
cap.name = "caption_chart_revenue"
cap.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
cap.text_frame.text = "Revenue by quarter ($M)"
# cap bottom ≈ 1.5 + 0.5 = 2.0in, chart top = 2.1in → within 0.1in ✓
```

---

## Units & Coordinates Quick Reference

```python
from pptx.util import Inches, Pt, Emu, Cm
Inches(1)   # 914400 EMU
Pt(1)       # 12700 EMU
Cm(1)       # 360000 EMU
```

---

## Common Pitfalls (python-pptx)

These cause bugs, corruption, or silent `check_pptx` failures.

1. **Colors are `RGBColor(0xRR, 0xGG, 0xBB)`** — never a `"#hex"` string.
2. **Positions/sizes need `Length` objects** (`Inches`/`Pt`/`Emu`), never bare floats.
3. **Always set `tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`** — so text isn't clipped when rendered. (It does not make overflow visible to `check_pptx`; size boxes generously and catch overflow with a visual render.)
4. **Always set `shape.name`** — unnamed shapes report as `'shape'`, making problems unreadable.
5. **Lines are connectors**, not `add_shape` — use `add_connector(MSO_CONNECTOR.STRAIGHT, ...)`.
6. **Four features need XML workarounds**: transparency/alpha, gradients, true bullet glyphs in plain textboxes, and reading an inherited background. For full-bleed color use the public `slide.background.fill` instead.
7. **Charts and tables need a caption** within 0.1in above them, or `element-boxing` fails.
8. **Two filled boxes that overlap are always flagged** — don't stack content boxes on a filled panel; use the true `<p:bg>` background for full-bleed color instead.
