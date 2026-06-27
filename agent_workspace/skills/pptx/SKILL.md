---
name: pptx
description: 'Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from a .pptx; editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions "deck," "slides," "presentation," or references a .pptx filename, regardless of what they plan to do with the content afterward. This skill builds decks with python-pptx and validates layout with check_pptx.'
---

# PPTX Skill (python-pptx + check_pptx)

## Quick Reference

| Task                    | Guide                                                            |
| ----------------------- | ---------------------------------------------------------------- |
| Read / extract text     | `python -m markitdown deck.pptx`                                 |
| Create from scratch     | Read [python-pptx.md](python-pptx.md)                            |
| Edit an existing deck    | Read [editing.md](editing.md)                                    |
| Design rules (read first) | Read [design.md](design.md) **before building**               |
| Name shapes correctly   | Read [naming.md](naming.md)                                      |
| Validate layout         | `python main.py deck.pptx` (or `from harness import check_pptx`) |

This skill builds with the **python-pptx** library and validates with **`check_pptx`** — the geometric layout validator in this project at `harness/tools.py`.

---

## Reading Content

```bash
python -m markitdown deck.pptx
```

Use this to extract text, check order, and find leftover placeholder content.

---

## Creating from Scratch

**Read [python-pptx.md](python-pptx.md) for the full build API.** One non-negotiable rule, because it is what `check_pptx` depends on:

1. **Every shape:** a meaningful `shape.name` (see [naming.md](naming.md)).

**Read [design.md](design.md) before you build** — it is the single source for palettes, typography, layout patterns, spacing constants, and the grader conventions a deck must satisfy. (Design guidance lives there, not here, so there's one home for it.)

---

## QA (Required) — via `check_pptx`

**Assume there are problems. Your job is to find them.** Your first render is almost never correct.

### Size text boxes generously — `auto_size` won't save you

**The single most important thing to understand about `check_pptx`:** it reads only **stored** geometry — the `left/top/width/height` written into the file (`shape.left`, etc.). Setting `tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT` writes an autofit _flag_ (`<a:spAutoFit/>`) into the XML; it does **not** rewrite the stored width/height. The box only physically grows when a renderer (PowerPoint/LibreOffice) opens and reflows it — and that grown size is never written back to the stored geometry (verified: even a LibreOffice round-trip leaves the stored height unchanged).

> **Therefore `auto_size` has no effect on what `check_pptx` sees — don't rely on it to fix overflow.** `check_pptx` cannot detect text overflowing inside a box from the autofit flag; its docstring says exactly this ("cannot detect text overflowing inside a box; python-pptx does not reflow text").

What this means in practice: **size your text boxes generously from the start**, and lean on `check_pptx`'s *estimated* overflow check (`typography-text-fits-its-box`) as the automated overflow signal — `check_pptx` itself is a geometry gate for the boxes you declared, not a text-fit checker. **There is no vision model in this toolkit**, so you cannot fall back on "just render it and look" (see Known blind spots below) — generous sizing + the estimated check are your real defenses.

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
  "problems": [
    { "slide": 2,
      "message": "'card_revenue' and 'card_cost' overlap by 0.31 in^2",
      "shapes": ["card_revenue", "card_cost"], "boxes_in": [ {...}, {...} ] }
  ],
  "notes": []
}
```

For each problem: `message` describes the defect, `slide` is 1-based, `shapes` names the offenders (this is why naming matters — see [naming.md](naming.md)), and `boxes_in` gives their boxes in inches so you know exactly what to move. `ok` is `true` only when there are **zero** problems.

### What it checks and how to fix each

Single tolerance throughout: **`TOL = 0.1in`**. Two coordinates closer than that count as "the same." The `message` on each problem tells you which row applies.

| When                                                                          | Fix                                                                                            |
| ------------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------- |
| a shape is off-slide or within 0.1in of any edge                               | move it inward; keep ~0.5in clear                                                              |
| any two shapes' boxes intersect                                                | separate them; don't stack content on filled panels (use `<p:bg>` for full-bleed color)        |
| a slide uses >2 left edges, or ≥2 slides deviate from the deck's dominant left | align primary elements to one shared left edge across the deck                                 |
| peer boxes in a row have unequal widths or non-aligned tops                    | give row peers identical width and top                                                         |
| gaps in a row/column of 3+ vary                                                | make the gaps equal                                                                            |
| a chart/table has no caption within 0.1in above it                             | add a caption textbox whose bottom is within 0.1in of the chart/table top, overlapping it in x |
| the title-top-y or footer-top-y varies across slides                           | put the title (and footer) at the same y on every slide                                        |

### Known blind spots

`check_pptx` is geometry-only, reading the boxes you declared. It will **not** catch:

- **In-box text overflow** — text spilling past the bottom/edge of its box. The autofit flag has no effect on what `check_pptx` reads (see above). The `typography-text-fits-its-box` check *estimates* this (trust a finding; its absence means "probably fine, still a blind spot"); for overflow into empty space below a box, you cannot detect it here — so size generously.
- **The true slide background** — full-bleed color set with `slide.background.fill` lives in `<p:bg>`, is not a shape, and is never walked. (Good: a full-bleed background can't trip overlap/margin checks.)
- **Decorative `kind="other"` shapes** — a filled rectangle with no text frame is classified `other`. It is **skipped** by alignment/left-margin/gutter checks but **still counted** by overlap and outer-margin. So a background band can collide or bleed off the edge and _will_ be flagged, but won't be forced to align to the grid.
- **Color, contrast, and font choices** — geometry says nothing about these. Follow `design.md`'s palette and WCAG ≥4.5:1 contrast rules **at authoring time**, because nothing here verifies them after the fact.

> **No vision model.** This toolkit has **no vision capability** — the agent cannot see a rendered slide. There is no "render it and inspect the image (yourself or with a subagent)" step available here. Your defenses against these blind spots are entirely up-front: size boxes generously, trust the estimated overflow check, and apply `design.md`'s color/contrast rules as you write the code.

You can still render to images as **output for a human to eyeball** (the agent cannot read them back):

```bash
soffice --headless --convert-to pdf deck.pptx
pdftoppm -jpeg -r 150 deck.pdf slide      # -> slide-01.jpg, slide-02.jpg, ...
```

`check_pptx` (geometry) plus the estimated overflow check are the agent's automated gate; the render is a human's appearance check, not the agent's.

### Verification loop

1. Generate the deck.
2. Run `python main.py deck.pptx`.
3. If `ok` is `false`, read each problem's `shapes` and `boxes_in`, fix those named shapes.
4. Re-run. One fix often surfaces another (a moved box now overlaps something else).
5. Repeat until `ok: true` **with no `typography-text-fits-its-box` finding** — if that estimated-overflow check fires, grow the box and/or push the shape below it down, then re-run.
6. **When the deck is done** (`ok: true`, no overflow finding), **call the `finish_deck` tool** with the deck path and a short `log` (palette/motif, iterations, final score) to persist the deck and the run log. (This tool is available in the agent-run harness; if it isn't present, you're not in that harness and can skip it.)

**Success for the agent = `check_pptx` returns `ok: true` AND no estimated-overflow finding.** Because there is no vision model here, overflow-into-empty-space, contrast, and color stay blind spots — a **human** confirms those on the optional render; the agent cannot. Prevent them up front via generous sizing and `design.md`'s color/contrast rules.

---

## Dependencies

- `python-pptx>=1.0.2` — building and validating decks (already in `pyproject.toml`).
- `check_pptx` — `harness/tools.py` in this project; run via `python main.py deck.pptx`.
- `markitdown[pptx]` — text extraction (optional).
- LibreOffice (`soffice`) + Poppler (`pdftoppm`) — optional, only to render images for a human to eyeball (the agent has no vision model and cannot read them).

