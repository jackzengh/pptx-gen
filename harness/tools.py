from __future__ import annotations

from dataclasses import dataclass, field
from itertools import combinations

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn

from .text_metrics import measure_text_emu

EMU_PER_INCH = 914400
TOL = 0.1
# Reserved outer margin (inches) for the outer-margin-frame check — matches the
# rubric grader's 0.3in band and standard IB convention. Distinct from TOL,
# which is the fine alignment/overlap tolerance.
MARGIN_BAND = 0.3
CAPTION_GAP = 0.5
TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
PRIMARY_KINDS = {"title", "body", "chart", "table", "picture", "text"}

# --- Shape roles by name prefix --------------------------------------------
# FOUR semantic roles, each mapping to exactly ONE exemption behavior. Name a
# shape `<role>_<descriptor>` (e.g. background_sidebar, chartpart_wf_total,
# overlay_image_credit, meta_pagenum_3) to opt it out of the matching checks.
# Ordinary content (title/body/card_*/chart_*/table_*) stays fully checked.
#
#   background_*  full-bleed panels & image fields — meant to touch the edge and
#                 host text on top. Exempt from outer-margin + being an overlap
#                 victim.
#   chartpart_*   any piece of a HAND-DRAWN chart (waterfall bars, comparison
#                 bars, scatter points & labels, brackets, multiplier row, axis
#                 labels). The chart is one unit: its parts may vary in
#                 width/gutter and overlap each other. Exempt from grid + their
#                 mutual overlaps.
#   overlay_*     a label that sits ON something (a caption on an image, a tile
#                 badge, an on-chart annotation). Exempt from overlap.
#   meta_*        non-grid chrome (eyebrow, footer, source, page number, nav,
#                 rules, legend, icon-row sub-headers). Exempt from grid checks.
#
# The bare descriptors below are legacy aliases kept so older decks still pass;
# new decks should use the four role prefixes above.
_DECOR_PREFIXES = (
    "background", "image", "bg", "sidebar", "panel", "field",  # legacy
)
# Chart-internal pieces (container + children, drawn by hand).
_PLOT_PREFIXES = ("chartpart", "plot")  # container region
_PLOT_CHILD_PREFIXES = (
    "chartpart",
    "pt", "ptlbl", "bub", "bublbl", "annot", "marker", "dot",  # legacy
)
_MANUAL_CHART_PREFIXES = (
    "chartpart",
    "wf", "cmp", "bar", "step", "waterfall", "bridge",         # legacy
    "mult", "multlabel", "catlabel", "axis", "grpbar",
)
# Labels placed on a panel / image / tile / chart.
_OVERLAY_PREFIXES = (
    "overlay",
    "caption", "tile", "tilenum", "annot", "bracket", "tick",  # legacy
)
# Non-grid chrome.
_META_PREFIXES = (
    "meta",
    "eyebrow", "title_rule", "rule", "nav", "logo", "footer", "source",  # legacy
    "pagenum", "pageno", "divider", "vrule", "legend", "icon", "indhdr",
    "inddesc", "iconlbl", "head", "subhead", "label", "iconlabel",
    "callout", "operand", "opcap",
)


def _has_role(name: str, prefixes: tuple) -> bool:
    n = (name or "").lower()
    return any(n == p or n.startswith(p + "_") or n.endswith("_" + p) for p in prefixes)


def _is_decor(b: "Box") -> bool:
    """True for full-bleed/panel shapes meant to bleed and host overlaid text."""
    return _has_role(b.name, _DECOR_PREFIXES)


def _contained_in(child: "Box", parent: "Box", pad: float = TOL) -> bool:
    return (child.left >= parent.left - pad and child.top >= parent.top - pad
            and child.right <= parent.right + pad and child.bottom <= parent.bottom + pad)


# Everything that is NOT a main-content grid block (excluded from left-margin /
# row-width / gutter / column-alignment checks): all four roles plus content
# that lives inside a background panel (handled separately in _grid_boxes).
_NON_GRID_PREFIXES = (
    _DECOR_PREFIXES + _OVERLAY_PREFIXES + _PLOT_PREFIXES
    + _PLOT_CHILD_PREFIXES + _MANUAL_CHART_PREFIXES + _META_PREFIXES
)


def _grid_boxes(slide: "SlideInfo") -> list:
    """Main-content-grid shapes: exclude decor panels, anything inside a panel,
    eyebrows, overlays, plot internals, rules, footers, and sources.
    """
    panels = [b for b in slide.boxes if _is_decor(b)]
    out = []
    for b in slide.boxes:
        if _has_role(b.name, _NON_GRID_PREFIXES):
            continue
        if any(_contained_in(b, p) for p in panels):  # sidebar content, etc.
            continue
        out.append(b)
    return out

# --- Text-overflow estimation constants -----------------------------------
# A proportional glyph's average advance width is ~0.5-0.6 em. Titles here use
# bold/serif faces that run wider, so 0.58 em is the calibrated value that makes
# the line-count estimate match what PowerPoint actually renders (validated
# against rendered slides). LINE_SPACING is the multiple of font size each line
# occupies vertically (single-spaced text renders at ~1.2x its point size).
CHAR_WIDTH_EM = 0.58
LINE_SPACING = 1.2
PT_PER_INCH = 72.0
# Fallback point sizes when a run leaves its size unset (inherited from theme).
DEFAULT_PT = {"title": 36.0, "body": 14.0, "text": 14.0}


def emu_to_in(value: int) -> float:
    """Convert a python-pptx length or raw EMU value to inches."""
    inches = value.inches if hasattr(value, "inches") else value / EMU_PER_INCH
    return round(inches, 3)


# --------------------------------------------------------------------------- #
# Shape collection
# --------------------------------------------------------------------------- #

@dataclass
class Box:
    """One leaf shape, with all geometry already in inches (slide coords)."""

    name: str
    kind: str          # title | body | chart | table | picture | text | other
    left: float
    top: float
    width: float
    height: float
    text: str = ""
    # Per-paragraph (char_count, font_size_pt) for the text frame, plus the four
    # internal text insets (inches). Populated only for shapes that carry text;
    # used by check_text_overflow to estimate how tall the text renders. None
    # font sizes are filled with a slide/kind default at estimation time.
    paras: list = field(default_factory=list)
    text_insets: tuple = (0.1, 0.1, 0.05, 0.05)  # (left, right, top, bottom)

    @property
    def right(self) -> float:
        return round(self.left + self.width, 3)

    @property
    def bottom(self) -> float:
        return round(self.top + self.height, 3)

    def as_dict(self) -> dict:
        return {
            "left": self.left,
            "top": self.top,
            "width": self.width,
            "height": self.height,
        }


@dataclass
class SlideInfo:
    index: int                       # 1-based, for human-facing messages
    boxes: list[Box] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


@dataclass
class DeckContext:
    slide_w: float                   # inches
    slide_h: float                   # inches
    slides: list[SlideInfo]


def _kind_of(shape, title_shape) -> str:
    """Classify a shape into one of our coarse kinds."""
    if shape is title_shape:
        return "title"
    # has_table / has_chart are safe on any shape and never raise.
    if getattr(shape, "has_table", False):
        return "table"
    if getattr(shape, "has_chart", False):
        return "chart"
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if shape.is_placeholder:
        ph_type = shape.placeholder_format.type
        if ph_type in TITLE_TYPES:
            return "title"
        if ph_type == PP_PLACEHOLDER.BODY:
            return "body"
    # Only a shape carrying actual text counts as a "text" primary block. An
    # empty text frame (a decorative rectangle, a motif chip, a fill-only card)
    # is "other": still checked for overlap/margins, but excluded from the
    # left-margin / row / gutter alignment checks meant for content blocks.
    if (getattr(shape, "has_text_frame", False) and shape.has_text_frame
            and shape.text_frame.text.strip()):
        return "text"
    return "other"


def _group_transform(group):
    """Return the affine map (child coords -> parent coords) for a group.

    Reads off/ext (group box in parent coords) and chOff/chExt (child coord
    space) from the group's xfrm. Returns a closure mapping a child bbox
    (in EMU) to its bbox in the parent's coordinate space (in EMU).
    """
    xfrm = group._element.grpSpPr.xfrm
    ox, oy = xfrm.off.x, xfrm.off.y
    ecx, ecy = xfrm.ext.cx, xfrm.ext.cy
    cox, coy = xfrm.chOff.x, xfrm.chOff.y
    chcx, chcy = xfrm.chExt.cx, xfrm.chExt.cy

    # Guard degenerate child extents (would divide by zero).
    sx = (ecx / chcx) if chcx else 1.0
    sy = (ecy / chcy) if chcy else 1.0

    def to_parent(left, top, width, height):
        return (
            ox + (left - cox) * sx,
            oy + (top - coy) * sy,
            width * sx,
            height * sy,
        )

    return to_parent


def _text_metrics(shape):
    """Per-paragraph (char_count, font_size_pt) and the text-frame insets (in).

    Returns ([], default_insets) for shapes without a text frame. ``font_size_pt``
    is None when the run inherits its size from the theme; the overflow estimator
    substitutes a kind-based default for those. Insets default to PowerPoint's
    standard 0.1 in left/right and 0.05 in top/bottom when not set on the box.
    """
    paras = []
    insets = (0.1, 0.1, 0.05, 0.05)
    if not (getattr(shape, "has_text_frame", False) and shape.has_text_frame):
        return paras, insets
    tf = shape.text_frame
    for p in tf.paragraphs:
        size = None
        bold = italic = False
        name = None
        for run in p.runs:
            if run.font.size is not None:
                size = run.font.size.pt
            if run.font.bold:
                bold = True
            if run.font.italic:
                italic = True
            if run.font.name:
                name = run.font.name
        # (text, font_size_pt|None, bold, italic, font_name|None) — text and
        # style are needed for real-font wrap measurement (see text_metrics).
        paras.append((p.text, size, bold, italic, name))
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    if bodyPr is not None:
        def inset(attr, default):
            v = bodyPr.get(attr)
            return (int(v) / EMU_PER_INCH) if v is not None else default
        insets = (
            inset("lIns", 0.1), inset("rIns", 0.1),
            inset("tIns", 0.05), inset("bIns", 0.05),
        )
    return paras, insets


def _walk(shape, title_shape, transforms, out: list[Box], notes: list[str]):
    """Recursively collect leaf shapes, applying group transforms.

    ``transforms`` is the chain of group maps from outermost to innermost; each
    leaf's EMU bbox is pushed back through every map (innermost first) to reach
    absolute slide coordinates.
    """
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        to_parent = _group_transform(shape)
        for child in shape.shapes:
            _walk(child, None, transforms + [to_parent], out, notes)
        return

    left, top = shape.left, shape.top
    width, height = shape.width, shape.height
    if None in (left, top, width, height):
        # Inherited-geometry placeholder (no explicit position). Skip but note.
        notes.append(f"{shape.name or 'shape'}: inherited geometry (skipped)")
        return

    left, top, width, height = int(left), int(top), int(width), int(height)
    # Apply group maps from innermost outward to reach slide coordinates.
    for to_parent in reversed(transforms):
        left, top, width, height = to_parent(left, top, width, height)

    paras, insets = _text_metrics(shape)
    out.append(
        Box(
            name=shape.name or "shape",
            kind=_kind_of(shape, title_shape),
            left=emu_to_in(left),
            top=emu_to_in(top),
            width=emu_to_in(width),
            height=emu_to_in(height),
            text=(
                shape.text_frame.text.strip()
                if getattr(shape, "has_text_frame", False) and shape.has_text_frame
                else ""
            ),
            paras=paras,
            text_insets=insets,
        )
    )


def collect_shapes(slide, index: int) -> SlideInfo:
    """Flatten a slide into leaf Boxes in absolute slide inches."""
    info = SlideInfo(index=index + 1)
    title_shape = slide.shapes.title  # may be None
    for shape in slide.shapes:
        _walk(shape, title_shape, [], info.boxes, info.notes)
    return info


# --------------------------------------------------------------------------- #
# Problem helpers
# --------------------------------------------------------------------------- #

def _problem(criterion, slide, message, shapes=None, boxes=None) -> dict:
    return {
        "criterion": criterion,
        "slide": slide,
        "message": message,
        "shapes": shapes or [],
        "boxes_in": boxes or [],
    }


def _intersection_area(a: Box, b: Box) -> float:
    dx = min(a.right, b.right) - max(a.left, b.left)
    dy = min(a.bottom, b.bottom) - max(a.top, b.top)
    if dx > 0 and dy > 0:
        return dx * dy
    return 0.0


def _effective_bottom(b: Box) -> float:
    """The box's real ink bottom: rendered text height for a text box (which may
    be SHORTER than its declared box when auto_size grows the frame), else the
    declared bottom. Two stacked auto-size text boxes whose declared boxes touch
    but whose actual text doesn't should not count as overlapping.
    """
    if b.paras:
        est = _estimated_text_height(b)
        if est is not None:
            return round(b.top + est, 3)
    return b.bottom


def _ink_intersection_area(a: Box, b: Box) -> float:
    """Intersection using each box's effective (text-aware) bottom."""
    dx = min(a.right, b.right) - max(a.left, b.left)
    dy = min(_effective_bottom(a), _effective_bottom(b)) - max(a.top, b.top)
    if dx > 0 and dy > 0:
        return dx * dy
    return 0.0


# --------------------------------------------------------------------------- #
# Checks (one per rubric criterion)
# --------------------------------------------------------------------------- #

def check_outer_margin_frame(ctx: DeckContext) -> list[dict]:
    crit = "layout-and-alignment-outer-margin-frame"
    out = []
    for s in ctx.slides:
        for b in s.boxes:
            # Full-bleed panels / image fields are meant to touch the edge, as
            # are captions anchored to an image's bottom-left corner.
            if _is_decor(b) or _has_role(b.name, ("caption",)):
                continue
            off_slide = (
                b.left < 0 or b.top < 0
                or b.right > ctx.slide_w or b.bottom > ctx.slide_h
            )
            # Reserve a 0.3in margin on all edges (standard IB convention, and
            # what the rubric grader enforces) so page numbers / footers / any
            # content stay clear of the slide edge. Larger than TOL on purpose.
            in_band = (
                b.left < MARGIN_BAND or b.top < MARGIN_BAND
                or b.right > ctx.slide_w - MARGIN_BAND
                or b.bottom > ctx.slide_h - MARGIN_BAND
            )
            if off_slide:
                out.append(_problem(
                    crit, s.index,
                    f"{b.name!r} bleeds off the slide",
                    [b.name], [b.as_dict()],
                ))
            elif in_band:
                out.append(_problem(
                    crit, s.index,
                    f"{b.name!r} intrudes into the {MARGIN_BAND} in outer margin band",
                    [b.name], [b.as_dict()],
                ))
    return out


def check_no_overlap(ctx: DeckContext) -> list[dict]:
    crit = "layout-and-alignment-no-overlap-or-collision"
    out = []
    for s in ctx.slides:
        plots = [b for b in s.boxes if _has_role(b.name, _PLOT_PREFIXES)]
        for a, b in combinations(s.boxes, 2):
            # Use text-aware (rendered) bounds so two stacked auto-size text
            # boxes aren't flagged just because their declared boxes touch.
            area = _ink_intersection_area(a, b)
            if area <= 0:
                continue
            # Intentional overlaps in BCG-style layouts:
            # 1. text/labels sitting ON a full-bleed panel, image field, image
            #    caption, or a contents-tile (badge + caption on an image tile).
            if _is_decor(a) or _is_decor(b):
                continue
            if _has_role(a.name, _OVERLAY_PREFIXES) or _has_role(b.name, _OVERLAY_PREFIXES):
                continue
            # manual chart components (waterfall/comparison bars + their labels)
            if (_has_role(a.name, _MANUAL_CHART_PREFIXES)
                    and _has_role(b.name, _MANUAL_CHART_PREFIXES)):
                continue
            # 2. scatter/bubble markers, labels, and on-chart annotations that
            #    live inside a plot region — the chart is one visual unit.
            ab_plot_children = (
                _has_role(a.name, _PLOT_CHILD_PREFIXES + _PLOT_PREFIXES)
                and _has_role(b.name, _PLOT_CHILD_PREFIXES + _PLOT_PREFIXES)
            )
            if ab_plot_children:
                continue
            if any(
                (_contained_in(a, p) or _has_role(a.name, _PLOT_CHILD_PREFIXES))
                and (_contained_in(b, p) or _has_role(b.name, _PLOT_CHILD_PREFIXES))
                for p in plots
            ):
                continue
            out.append(_problem(
                crit, s.index,
                f"{a.name!r} and {b.name!r} overlap by "
                f"{round(area, 3)} in^2",
                [a.name, b.name], [a.as_dict(), b.as_dict()],
            ))
    return out


def check_left_margin_consistency(ctx: DeckContext) -> list[dict]:
    crit = "layout-and-alignment-left-margin-consistency"
    out = []

    # Per-slide: dominant left edge = the most common block left edge.
    # We count *block* left edges, not every shape's left: peer columns within a
    # row share one block (the row), so only the row's leftmost edge counts.
    # Otherwise a legitimate 3-column layout would read as 3 distinct margins.
    slide_lefts = {}
    for s in ctx.slides:
        lefts = _block_left_edges(_grid_boxes(s))
        if not lefts:
            continue
        # Cluster lefts within TOL; dominant cluster's representative value.
        clusters = _cluster(lefts)
        # Distinct clusters per slide should be <= 2.
        if len(clusters) > 2:
            out.append(_problem(
                crit, s.index,
                f"slide uses {len(clusters)} distinct block left edges "
                f"(expected at most 2)",
            ))
        dominant = max(clusters, key=lambda c: len(c))
        slide_lefts[s.index] = sum(dominant) / len(dominant)

    if len(slide_lefts) >= 2:
        # Deck dominant left = most common slide-dominant value (within TOL).
        values = list(slide_lefts.values())
        deck_clusters = _cluster(values)
        deck_dominant_cluster = max(deck_clusters, key=lambda c: len(c))
        deck_left = sum(deck_dominant_cluster) / len(deck_dominant_cluster)
        deviating = [
            idx for idx, v in slide_lefts.items()
            if abs(v - deck_left) > TOL
        ]
        if len(deviating) >= 2:
            out.append(_problem(
                crit, deviating[0],
                f"{len(deviating)} slides (e.g. {sorted(deviating)}) use a left "
                f"edge differing by more than {TOL} in from the deck's dominant "
                f"left edge of {round(deck_left, 3)} in",
            ))
    return out


def check_edge_and_grid_alignment(ctx: DeckContext) -> list[dict]:
    crit = "layout-and-alignment-edge-and-grid-alignment"
    out = []
    for s in ctx.slides:
        for row in _rows(_grid_boxes(s)):
            if len(row) < 2:
                continue
            widths = [b.width for b in row]
            tops = [b.top for b in row]
            if max(widths) - min(widths) > TOL:
                out.append(_problem(
                    crit, s.index,
                    f"peer boxes in a row have unequal widths "
                    f"({round(min(widths), 3)}..{round(max(widths), 3)} in)",
                    [b.name for b in row], [b.as_dict() for b in row],
                ))
            if max(tops) - min(tops) > TOL:
                out.append(_problem(
                    crit, s.index,
                    f"peer boxes in a row are not top-aligned "
                    f"({round(min(tops), 3)}..{round(max(tops), 3)} in)",
                    [b.name for b in row], [b.as_dict() for b in row],
                ))
    return out


def check_consistent_gutters(ctx: DeckContext) -> list[dict]:
    crit = "layout-and-alignment-consistent-gutters"
    out = []
    for s in ctx.slides:
        # Horizontal gutters within rows of 3+.
        for row in _rows(_grid_boxes(s)):
            if len(row) < 3:
                continue
            ordered = sorted(row, key=lambda b: b.left)
            gaps = [
                ordered[i + 1].left - ordered[i].right
                for i in range(len(ordered) - 1)
            ]
            if gaps and max(gaps) - min(gaps) > TOL:
                out.append(_problem(
                    crit, s.index,
                    f"horizontal gutters in a row of {len(row)} vary "
                    f"({round(min(gaps), 3)}..{round(max(gaps), 3)} in)",
                    [b.name for b in ordered],
                    [b.as_dict() for b in ordered],
                ))
        # Vertical gutters within columns of 3+ genuine stacked peers. A title
        # and a footer/source line are distinct roles, not list peers, so a
        # title -> body -> footer stack is NOT expected to have even gaps; drop
        # titles and bottom-band (footer/source) boxes before measuring.
        for col in _columns(_grid_boxes(s)):
            peers = [
                b for b in col
                if b.kind != "title" and b.bottom < ctx.slide_h - 0.6
            ]
            if len(peers) < 3:
                continue
            ordered = sorted(peers, key=lambda b: b.top)
            gaps = [
                ordered[i + 1].top - ordered[i].bottom
                for i in range(len(ordered) - 1)
            ]
            if gaps and max(gaps) - min(gaps) > TOL:
                out.append(_problem(
                    crit, s.index,
                    f"vertical gutters in a column of {len(col)} vary "
                    f"({round(min(gaps), 3)}..{round(max(gaps), 3)} in)",
                    [b.name for b in ordered],
                    [b.as_dict() for b in ordered],
                ))
    return out


def check_element_boxing(ctx: DeckContext) -> list[dict]:
    crit = "layout-and-alignment-element-boxing"
    out = []
    for s in ctx.slides:
        captions = [b for b in s.boxes if b.kind in ("text", "title") and b.text]
        for el in s.boxes:
            if el.kind not in ("chart", "table"):
                continue
            has_caption = any(
                # caption sits above the element, within CAPTION_GAP, x-overlapping
                (el.top - cap.bottom) >= -TOL
                and (el.top - cap.bottom) <= CAPTION_GAP
                and min(cap.right, el.right) - max(cap.left, el.left) > 0
                for cap in captions
            )
            if not has_caption:
                out.append(_problem(
                    crit, s.index,
                    f"{el.kind} {el.name!r} has no caption within "
                    f"{CAPTION_GAP} in above it",
                    [el.name], [el.as_dict()],
                ))
    return out


def _estimated_text_height(b: Box) -> float | None:
    """Estimate how tall b's text renders, in inches, accounting for wrapping.

    Measures with real font metrics (see text_metrics.measure_text_emu): the
    bundled Liberation faces are metric-compatible with the common Office/web
    fonts, so greedy word-wrap against the actual glyph advances yields the same
    line counts PowerPoint produces — far more accurate than an avg-char-width
    guess. A conservative WRAP_FILL_FACTOR biases near-full lines to wrap, so a
    blind model's text is flagged rather than silently overflowing.

    This is the one defect python-pptx geometry alone cannot see: a box whose
    declared height is fine but whose text reflows past it (and, per
    check_text_overflow, into the shape below). Returns None when the box carries
    no text or has no usable width.
    """
    if not b.paras:
        return None
    # An entirely-empty text frame (e.g. an oval/rectangle marker that happens to
    # carry a blank text frame) renders no text — don't estimate a phantom line.
    if all(not (txt or "").strip() for txt, *_ in b.paras):
        return None
    lI, rI, tI, bI = b.text_insets
    usable_w_in = b.width - lI - rI
    if usable_w_in <= 0:
        return None
    default_pt = DEFAULT_PT.get(b.kind, 14.0)

    total = tI + bI
    box_w_emu = int(usable_w_in * EMU_PER_INCH)
    for text, size, bold, italic, name in b.paras:
        pt = size if size is not None else default_pt
        _, h_emu = measure_text_emu(
            text or "",
            box_w_emu,
            pt,
            LINE_SPACING,
            font_family=name or "Arial",
            bold=bool(bold),
            italic=bool(italic),
        )
        total += h_emu / EMU_PER_INCH
    return total


def check_text_overflow(ctx: DeckContext) -> list[dict]:
    """Flag text that reflows past its box AND collides with the shape below it.

    python-pptx exposes box geometry but never reflows text, so a title sized for
    one line that actually wraps to three reads as 'fine' geometrically while
    visibly crashing into the content beneath it. We estimate the rendered text
    height (see _estimated_text_height); if the text bottom clears the box bottom
    by more than TOL we then check whether that overspill region intersects any
    other shape lower on the slide. Reporting only collisions (not every box that
    is a hair too short) keeps this matched to the defect a reader actually sees.
    """
    crit = "typography-text-fits-its-box"
    out = []
    for s in ctx.slides:
        for b in s.boxes:
            est = _estimated_text_height(b)
            if est is None:
                continue
            text_bottom = b.top + est
            if text_bottom <= b.bottom + TOL:
                continue  # text fits inside its own box
            # Overspill exists: does it land on a shape below?
            for other in s.boxes:
                if other is b:
                    continue
                # horizontal overlap with the text box
                if min(b.right, other.right) - max(b.left, other.left) <= 0:
                    continue
                # vertical overlap of the overspill band with the other shape
                overlap = (min(text_bottom, other.bottom)
                           - max(b.bottom, other.top))
                if overlap > TOL:
                    out.append(_problem(
                        crit, s.index,
                        f"{b.name!r} text overflows its box (≈{round(est, 2)} in "
                        f"tall vs {round(b.height, 2)} in) and collides with "
                        f"{other.name!r} below",
                        [b.name, other.name],
                        [b.as_dict(), other.as_dict()],
                    ))
                    break
    return out


def check_y_offset_variation(ctx: DeckContext) -> list[dict]:
    crit = "layout-and-alignment-vertical-rhythm"
    out = []
    title_tops = {}
    footer_tops = {}
    for s in ctx.slides:
        # Title rhythm only across grid (content) titles — eyebrow slides shift
        # the title down, and section/cover titles live in a sidebar.
        grid = _grid_boxes(s)
        panels = [b for b in s.boxes if _is_decor(b)]
        titles = [b for b in grid if b.kind == "title"]
        if titles:
            title_tops[s.index] = min(b.top for b in titles)
        # Footer/source rhythm: only the standard bottom-band source line, NOT a
        # footer placed inside a sidebar panel (cover/section slides).
        footers = [
            b for b in s.boxes
            if b.kind in ("text", "body") and b.top > ctx.slide_h - 1.0
            and not any(_contained_in(b, p) for p in panels)
        ]
        if footers:
            footer_tops[s.index] = max(footers, key=lambda b: b.top).top

    out += y_offset_variation(crit, "title", title_tops)
    out += y_offset_variation(crit, "footer/page-number", footer_tops)
    return out


# --------------------------------------------------------------------------- #
# Small geometry / clustering utilities
# --------------------------------------------------------------------------- #

def _cluster(values: list[float]) -> list[list[float]]:
    """Greedy 1-D clustering: values within TOL of a cluster's mean join it."""
    clusters: list[list[float]] = []
    for v in sorted(values):
        for c in clusters:
            if abs(v - (sum(c) / len(c))) <= TOL:
                c.append(v)
                break
        else:
            clusters.append([v])
    return clusters


# Titles span the slide and are not grid peers; exclude them from row/column
# peer detection so a title stacked above content isn't mistaken for a column.
PEER_KINDS = {"body", "chart", "table", "picture", "text"}


def _rows(boxes: list[Box]) -> list[list[Box]]:
    """Group boxes into peer rows.

    A peer row is 2+ boxes that share a top (within TOL) AND have similar
    heights (within TOL) -- i.e. a genuine multi-column band, not a title that
    happens to share a top with something.
    """
    peers = [b for b in boxes if b.kind in PEER_KINDS]
    rows: list[list[Box]] = []
    for b in sorted(peers, key=lambda x: x.top):
        for row in rows:
            if abs(b.top - row[0].top) <= TOL and abs(b.height - row[0].height) <= TOL:
                row.append(b)
                break
        else:
            rows.append([b])
    return [r for r in rows if len(r) >= 2]


def _columns(boxes: list[Box]) -> list[list[Box]]:
    """Group boxes into peer columns.

    A peer column is 2+ boxes that share a left (within TOL) AND have similar
    widths (within TOL) -- a genuine stacked column, not coincidentally
    left-aligned heterogeneous shapes (e.g. a title above a chart).
    """
    peers = [b for b in boxes if b.kind in PEER_KINDS]
    cols: list[list[Box]] = []
    for b in sorted(peers, key=lambda x: x.left):
        for col in cols:
            if abs(b.left - col[0].left) <= TOL and abs(b.width - col[0].width) <= TOL:
                col.append(b)
                break
        else:
            cols.append([b])
    return [c for c in cols if len(c) >= 2]


def _block_left_edges(boxes: list[Box]) -> list[float]:
    """Left edges of block-level elements, one per peer row.

    Peer columns within a row collapse to the row's leftmost edge, so a clean
    N-column layout contributes a single block left edge rather than N.
    """
    rows = _rows(boxes)
    row_members = {id(b) for row in rows for b in row}
    lefts = []
    # Each peer row contributes only its leftmost edge.
    for row in rows:
        lefts.append(min(b.left for b in row))
    # Also group standalone primaries that share a top into a "soft row": a
    # chart beside a table (different widths, so not strict peers) is still one
    # visual row and contributes only its leftmost edge — not one per element.
    standalone = [
        b for b in boxes
        if b.kind in PRIMARY_KINDS and id(b) not in row_members
    ]
    soft_rows: list[list[Box]] = []
    for b in sorted(standalone, key=lambda x: x.top):
        for r in soft_rows:
            if abs(b.top - r[0].top) <= TOL:
                r.append(b)
                break
        else:
            soft_rows.append([b])
    for r in soft_rows:
        lefts.append(min(b.left for b in r))
    return lefts


def y_offset_variation(crit, label, tops: dict[int, float]) -> list[dict]:
    if len(tops) < 2:
        return []
    values = list(tops.values())
    if max(values) - min(values) > TOL:
        slides = sorted(tops.keys())
        return [_problem(
            crit, slides[0],
            f"{label} top y-offset varies across slides "
            f"({round(min(values), 3)}..{round(max(values), 3)} in)",
        )]
    return []


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #

def check_pptx(path: str) -> dict:
    """Validate the layout of a saved .pptx and return structured problems.

    Returns a dict: {ok, problems, notes}. Each problem names the
    criterion, the 1-based slide number, the offending shapes, and their boxes
    (in inches), so the caller knows exactly what to move.

    NOTE: text overflow is ESTIMATED (python-pptx does not reflow text), via a
    calibrated wrap heuristic — check_text_overflow flags only overflow that
    collides with the shape below. Treat overflow findings as high-confidence but
    not pixel-exact; all other checks are exact geometry.
    """
    prs = Presentation(path)
    slides = [collect_shapes(s, i) for i, s in enumerate(prs.slides)]
    ctx = DeckContext(
        slide_w=emu_to_in(prs.slide_width),
        slide_h=emu_to_in(prs.slide_height),
        slides=slides,
    )

    problems: list[dict] = []
    problems += check_outer_margin_frame(ctx)
    problems += check_no_overlap(ctx)
    problems += check_left_margin_consistency(ctx)
    problems += check_edge_and_grid_alignment(ctx)
    problems += check_consistent_gutters(ctx)
    problems += check_element_boxing(ctx)
    problems += check_text_overflow(ctx)
    problems += check_y_offset_variation(ctx)

    notes = [
        {"slide": s.index, "notes": s.notes}
        for s in slides if s.notes
    ]
    return {
        "ok": not problems,
        "problems": problems,
        "notes": notes,
    }
