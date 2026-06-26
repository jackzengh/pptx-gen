---
name: pptx
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from a .pptx; editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions \"deck,\" \"slides,\" \"presentation,\" or references a .pptx filename, regardless of what they plan to do with the content afterward. This skill builds decks with python-pptx and validates layout with check_pptx."
---

# PPTX Skill (python-pptx + check_pptx)

## Quick Reference

| Task | Guide |
|------|-------|
| Read / extract text | `python -m markitdown deck.pptx` |
| Create from scratch | Read [python-pptx.md](python-pptx.md) |
| Name shapes correctly | Read [naming.md](naming.md) |
| Validate layout | `python main.py deck.pptx` (or `from harness import check_pptx`) |

This skill builds with the **python-pptx** library and validates with **`check_pptx`** — the geometric layout validator in this project at `harness/tools.py`.

---

## Reading Content

```bash
python -m markitdown deck.pptx
```

Use this to extract text, check order, and find leftover placeholder content.

---

## Creating from Scratch

**Read [python-pptx.md](python-pptx.md) for the full build API.** Two non-negotiable rules, because they are what `check_pptx` depends on:

1. **Every text frame:** `tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`.
2. **Every shape:** a meaningful `shape.name` (see [naming.md](naming.md)).

---

## Design Ideas

**Don't create boring slides.** Plain bullets on a white background won't impress anyone.

### Before starting

- **Pick a bold, content-informed palette.** If swapping your colors into a totally different deck would still "work," they aren't specific enough.
- **Dominance over equality.** One color dominates (60–70% of visual weight), with 1–2 supporting tones and one sharp accent. Never give all colors equal weight.
- **Dark/light contrast.** Dark backgrounds for title + conclusion, light for content ("sandwich"). Or commit to dark throughout for a premium feel.
- **Commit to a visual motif.** Pick ONE distinctive element and repeat it on every slide — rounded image frames, icons in colored circles, a thick single-side band.

### Color palettes

Don't default to generic blue. Use these as inspiration:

| Theme | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| Midnight Executive | `1E2761` | `CADCFC` | `FFFFFF` |
| Forest & Moss | `2C5F2D` | `97BC62` | `F5F5F5` |
| Coral Energy | `F96167` | `F9E795` | `2F3C7E` |
| Warm Terracotta | `B85042` | `E7E8D1` | `A7BEAE` |
| Ocean Gradient | `065A82` | `1C7293` | `21295C` |
| Charcoal Minimal | `36454F` | `F2F2F2` | `212121` |
| Teal Trust | `028090` | `00A896` | `02C39A` |
| Berry & Cream | `6D2E46` | `A26769` | `ECE2D0` |
| Cherry Bold | `990011` | `FCF6F5` | `2F3C7E` |

> In python-pptx, write these as `RGBColor(0x1E, 0x27, 0x61)` — never `"#1E2761"`.

### Layout & data display

**Every slide needs a visual element** — image, chart, icon, or shape. Text-only slides are forgettable.

- Two-column (text left, visual right) · icon + text rows · 2×2 / 2×3 grids · half-bleed image with overlay.
- Large stat callouts (60–72pt numbers, small labels) · before/after columns · numbered process flows.

### Typography

Choose an interesting font pairing — don't default to Arial.

| Element | Size |
|---------|------|
| Slide title | 36–44pt bold |
| Section header | 20–24pt bold |
| Body text | 14–16pt |
| Captions | 10–12pt muted |

Pairings: Georgia/Calibri · Arial Black/Arial · Cambria/Calibri · Trebuchet MS/Calibri · Palatino/Garamond.

### Spacing

- 0.5" minimum margins (the outer 0.1" is a hard fail — keep well clear).
- 0.3–0.5" between content blocks. Pick one gap and use it consistently.
- Leave breathing room — don't fill every inch.

### Avoid (common mistakes)

- **Don't repeat the same layout** — vary columns, cards, and callouts.
- **Don't center body text** — left-align paragraphs and lists; center only titles.
- **Don't skimp on size contrast** — titles 36pt+ vs 14–16pt body.
- **Don't default to blue** — pick colors that reflect the topic.
- **Don't mix spacing randomly** — choose 0.3" or 0.5" gaps and stick to it.
- **Don't create text-only slides.**
- **Don't forget text-box padding** — when aligning text to a shape edge, zero the frame margins (`tf.margin_left = Pt(0)`, etc.).
- **NEVER use accent lines under titles** — a hallmark of AI-generated slides. Use whitespace or background color instead.

---

## QA (Required) — via `check_pptx`

**Assume there are problems. Your job is to find them.** Your first render is almost never correct.

### Always set `auto_size` — and know exactly what it does and doesn't do

Set `tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT` on **every** text frame (rule 1). It is the correct authoring choice: in PowerPoint and LibreOffice the box grows to fit its text instead of clipping it, so your slides render without truncated content.

**But be clear about the limit — this is the single most important thing to understand about `check_pptx`:**

`check_pptx` reads only **stored** geometry — the `left/top/width/height` written into the file (`shape.left`, etc.). Setting `auto_size` writes an autofit *flag* (`<a:spAutoFit/>`) into the XML; it does **not** rewrite the stored width/height. The box only physically grows when a renderer (PowerPoint/LibreOffice) opens and reflows it — and that grown size is never written back to the stored geometry (verified: even a LibreOffice round-trip leaves the stored height unchanged).

> **Therefore `check_pptx` cannot detect text overflowing inside a box — full stop.** Its docstring says exactly this ("cannot detect text overflowing inside a box; python-pptx does not reflow text"). Do not assume `auto_size` makes overflow show up in the report. It does not.

What this means in practice: size your text boxes generously from the start, and **catch overflow with a visual render** (see Known blind spots below), not with `check_pptx`. `check_pptx` is a geometry gate for the boxes you declared — not a text-fit checker.

### Run it

```bash
python main.py deck.pptx      # prints the JSON report
```

Or in Python:

```python
from harness import check_pptx
report = check_pptx("deck.pptx")
```

The report is:

```json
{
  "ok": false,
  "summary": { "score": 34, "max_score": 42, "fail_count": 2, "by_criterion": {...} },
  "problems": [
    { "criterion": "...", "slide": 2, "severity": "fail",
      "message": "'card_revenue' and 'card_cost' overlap by 0.31 in^2",
      "shapes": ["card_revenue", "card_cost"], "boxes_in": [ {...}, {...} ] }
  ],
  "notes": []
}
```

For each problem: `slide` is 1-based, `shapes` names the offenders (this is why naming matters — see [naming.md](naming.md)), and `boxes_in` gives their boxes in inches so you know exactly what to move. `ok` is `true` only when there are **zero** problems.

### The 7 criteria and how to fix each

Single tolerance throughout: **`TOL = 0.1in`**. Two coordinates closer than that count as "the same."

Criterion ids below are shortened; in the JSON each is prefixed with `layout-and-alignment-` (e.g. `layout-and-alignment-no-overlap-or-collision`).

| Criterion (weight) | Fails when | Fix |
|---|---|---|
| `outer-margin-frame` (6) | a shape is off-slide or within 0.1in of any edge | move it inward; keep ~0.5in clear |
| `no-overlap-or-collision` (8) | any two shapes' boxes intersect | separate them; don't stack content on filled panels (use `<p:bg>` for full-bleed color) |
| `left-margin-consistency` (8) | a slide uses >2 left edges, or ≥2 slides deviate from the deck's dominant left | align primary elements to one shared left edge across the deck |
| `edge-and-grid-alignment` (6) | peer boxes in a row have unequal widths or non-aligned tops | give row peers identical width and top |
| `consistent-gutters` (5) | gaps in a row/column of 3+ vary | make the gaps equal |
| `element-boxing` (6) | a chart/table has no caption within 0.1in above it | add a caption textbox whose bottom is within 0.1in of the chart/table top, overlapping it in x |
| `vertical-rhythm` (3) | the title-top-y or footer-top-y varies across slides | put the title (and footer) at the same y on every slide |

### Known blind spots

`check_pptx` is geometry-only, reading the boxes you declared. It will **not** catch:

- **In-box text overflow** — text spilling past the bottom/edge of its box. `auto_size` does not change this (see above). **This is the main thing you must check by eye.**
- **The true slide background** — full-bleed color set with `slide.background.fill` lives in `<p:bg>`, is not a shape, and is never walked. (Good: a full-bleed background can't trip overlap/margin checks.)
- **Decorative `kind="other"` shapes** — a filled rectangle with no text frame is classified `other`. It is **skipped** by alignment/left-margin/gutter checks but **still counted** by overlap and outer-margin. So a background band can collide or bleed off the edge and *will* be flagged, but won't be forced to align to the grid.
- **Color, contrast, and font choices** — geometry says nothing about these.

Because text overflow is a real blind spot, **always do a visual render** as a second gate — render to images and inspect them (yourself or with a subagent) for overflow, clipping, contrast, and color:

```bash
soffice --headless --convert-to pdf deck.pptx
pdftoppm -jpeg -r 150 deck.pdf slide      # -> slide-01.jpg, slide-02.jpg, ...
```

`check_pptx` (geometry) and the visual render (overflow + appearance) are **both required** — they catch different classes of problem.

### Verification loop

1. Generate the deck.
2. Run `python main.py deck.pptx`.
3. If `ok` is `false`, read each problem's `shapes` and `boxes_in`, fix those named shapes.
4. Re-run. One fix often surfaces another (a moved box now overlaps something else).
5. Repeat until `ok: true`.
6. Render to images and inspect for text overflow, clipping, contrast, and color (geometry can't see these).

**Do not declare success until `check_pptx` returns `ok: true` AND a visual pass shows no overflow or appearance problems.**

---

## Dependencies

- `python-pptx>=1.0.2` — building and validating decks (already in `pyproject.toml`).
- `check_pptx` — `harness/tools.py` in this project; run via `python main.py deck.pptx`.
- `markitdown[pptx]` — text extraction (optional).
- LibreOffice (`soffice`) + Poppler (`pdftoppm`) — optional, for the secondary visual render.
