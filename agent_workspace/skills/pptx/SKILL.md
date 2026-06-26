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

## Content & IB Conventions

These are **content** rules, not geometry — `check_pptx` does not verify them, so build them in from the start. They are what separates a polished investment-banking-style deck from a generic one. Apply on every content slide (the cover and pure section dividers are exempt where noted).

- **Page numbers.** Put a page-number text box on **every content slide** (exclude only the cover and pure section dividers), numbering monotonically, in a consistent corner (e.g. bottom-right) at the same y-offset deck-wide.
- **Source notes.** **Every** slide with a chart, table, or hard number carries a footnote that starts with `Source:` or `Note:` (e.g. `Source: FY2021 10-K`). Same y-offset and small muted size deck-wide.
- **Numeric table columns are right- or decimal-aligned**, never left-aligned. Set `cell.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT` for every all-numeric column; label/text columns stay left-aligned.
- **Action titles.** Each content-slide title states a **claim with a verb** ("Margins expanded sharply…", "Mobile revenue nearly doubled…"), not a bare topic label ("Overview"). Aim for ≥5 words with a finite verb.
- **Consistent number formatting.** One format per metric deck-wide: fixed decimals (all multiples as `12.3x`), one currency scale per metric (don't mix `$3.37B` and `$3,373M` for the *same* metric), consistent unit symbols.
- **One message per slide.** Exactly one title-level headline; sub-heads ("Why it matters", "So what") are body-level, not a second competing title.
- **One title-case convention deck-wide** — pick sentence case OR title case for all titles and don't mix.

---

## QA (Required) — via `check_pptx`

**Assume there are problems. Your job is to find them.** Your first render is almost never correct.

### Always set `auto_size` — and know exactly what it does and doesn't do

Set `tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT` on **every** text frame (rule 1). It is the correct authoring choice: in PowerPoint and LibreOffice the box grows to fit its text instead of clipping it, so your slides render without truncated content.

**Understand what `auto_size` does in the stored file:**

`check_pptx` reads the **stored** geometry — the `left/top/width/height` written into the file (`shape.left`, etc.). Setting `auto_size` writes an autofit *flag* (`<a:spAutoFit/>`) into the XML; it does **not** rewrite the stored width/height. The box only physically grows when a renderer (PowerPoint/LibreOffice) opens and reflows it — and that grown size is never written back to the stored geometry (verified: even a LibreOffice round-trip leaves the stored height unchanged). So `auto_size` alone does not make the stored box match the rendered text.

**What `check_pptx` now does about this:** it **measures** how tall each text box's text actually renders — using real font metrics (the bundled fonts in `harness/fonts/`, accounting for the font family, size, bold/italic) to compute line wrapping — and flags `typography-text-fits-its-box` when a box's text overflows **and that overspill collides with the shape directly below it** (e.g. a one-line-tall title that wraps to three lines and crashes into the cards beneath). This is the most common real defect and `check_pptx` now catches it.

> The measurement is close to the renderer but not pixel-identical. It fires only on genuine collisions (it won't nag about a box that's a hair short with empty space below). Trust an overflow finding as a real defect to fix; a final visual render is still the confirmation, especially for wide serif faces at the wrap boundary.

What this means in practice: **size title boxes for the number of lines they will actually wrap to** (a long action-title in a 36–40pt face spans ~2–3 lines at full width; budget ~1.2×font-size per line) and leave clear vertical gap before the next element. If `check_pptx` reports an overflow collision, grow the box and/or push the lower shape down.

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

### The criteria and how to fix each

Single tolerance throughout: **`TOL = 0.1in`**. Two coordinates closer than that count as "the same."

Criterion ids below are shortened; in the JSON the layout ones are prefixed with `layout-and-alignment-` (e.g. `layout-and-alignment-no-overlap-or-collision`); the overflow one is `typography-text-fits-its-box`.

| Criterion (weight) | Fails when | Fix |
|---|---|---|
| `outer-margin-frame` (6) | a shape is off-slide or within 0.1in of any edge | move it inward; keep ~0.5in clear |
| `no-overlap-or-collision` (8) | any two shapes' boxes intersect | separate them; don't stack content on filled panels (use `<p:bg>` for full-bleed color) |
| `left-margin-consistency` (8) | a slide uses >2 left edges, or ≥2 slides deviate from the deck's dominant left | align primary elements to one shared left edge across the deck |
| `edge-and-grid-alignment` (6) | peer boxes in a row have unequal widths or non-aligned tops | give row peers identical width and top |
| `consistent-gutters` (5) | gaps in a row/column of 3+ vary | make the gaps equal |
| `element-boxing` (6) | a chart/table has no caption within 0.1in above it | add a caption textbox whose bottom is within 0.1in of the chart/table top, overlapping it in x |
| `text-fits-its-box` | a box's text is measured to wrap past its box AND overlap the shape below | grow the box to fit the wrapped lines, and/or move the lower shape down so the text clears it |
| `vertical-rhythm` (3) | the title-top-y or footer-top-y varies across slides | put the title (and footer) at the same y on every slide |

### Known blind spots

`check_pptx` is geometry-first (plus a real-font text-height measurement for overflow). It will **not** reliably catch:

- **In-box text overflow that does NOT hit another shape** — overflow is flagged only when the overspill collides with the shape below (the visible defect). Text that overflows into empty space, or a wide serif at the wrap boundary the measurement just misses, can still slip past. A visual render is the final word.
- **The true slide background** — full-bleed color set with `slide.background.fill` lives in `<p:bg>`, is not a shape, and is never walked. (Good: a full-bleed background can't trip overlap/margin checks.)
- **Decorative `kind="other"` shapes** — a filled rectangle with no text frame is classified `other`. It is **skipped** by alignment/left-margin/gutter checks but **still counted** by overlap and outer-margin. So a background band can collide or bleed off the edge and *will* be flagged, but won't be forced to align to the grid.
- **Color, contrast, and font choices** — geometry says nothing about these.
- **Content & IB conventions** — page numbers, source notes, numeric right-alignment, action titles (see "Content & IB Conventions" above). Not geometry; build them in from the start.

A visual render is still required as the second gate — render to images and inspect them (yourself or with a subagent) for overflow, clipping, contrast, and color:

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
