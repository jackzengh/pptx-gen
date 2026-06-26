# Task: Generate the Take-Two financial overview PowerPoint

Build a polished `.pptx` slide deck from the content spec and validate its layout until it passes.

## Inputs (all paths relative to the workspace root)

- `repo/spec.md` — the content spec. It defines **14 slides** for a Take-Two Interactive financial overview, each with a takeaway-sentence title and, per slide, some mix of: a chart, one or two tables, four "hero" stat cards, and a right-side bullet column. Read it fully before writing any code.
- `skills/pptx/SKILL.md` — the **PPTX skill**. Read it and the files it links (`python-pptx.md`, `naming.md`) before building. Follow it exactly.
- `repo/harness/` + `repo/main.py` — the `check_pptx` layout validator. Run it on your output.

## What to produce

- A file `repo/ttwo_overview.pptx` built with **python-pptx** (16:9), one slide per `##` section in `spec.md`, in order.
- Follow the skill's design guidance: a bold, content-informed palette (this is a games-publisher finance deck — pick something deliberate, not default blue); dark title/closing slides, light content slides; one repeated visual motif; every slide has a visual element; left-aligned body text; titles 36–44pt.

## Hard rules from the skill (these are what `check_pptx` enforces)

**Build through the reusable helpers** in `skills/pptx/python-pptx.md` ("Reusable helpers" section) — copy `textbox`, `paragraph`, `rectangle`, `connector`, `style_run`, and `_apply_bullet` into your `build_deck.py` and create every element through them. This guarantees the rules below for free: `textbox()` always sets `auto_size = SHAPE_TO_FIT_TEXT` and zero margins, and `paragraph(..., bullet=True, level=n)` produces real bullets via `p.level` + `<a:buChar>` with a hanging indent. Do NOT hand-roll `add_textbox`/`add_paragraph` calls that skip these.

1. **Name every shape** with a meaningful `shape.name` (e.g. `title`, `card_net_revenue`, `chart_revenue_trend`, `caption_chart_revenue`, `background_band`). See `skills/pptx/naming.md`.
2. **Every text frame:** `tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT` (the `textbox()` helper does this).
3. **Bullets:** use `paragraph(..., bullet=True, level=n)` — real glyphs driven by `p.level`, never a literal "• " typed into the text.
4. **Full-bleed backgrounds** use `slide.background.fill` (the true `<p:bg>`), NOT a shape.
5. **No overlaps**, keep ≥0.5in from slide edges, align primary elements to one shared left edge across the deck, give row peers equal width + equal top, even gutters in rows/columns of 3+, and **every chart and table needs a caption text box whose bottom is within 0.1in directly above it**, and keep the title at the same y on every slide.

## How to run things

This workspace is a local sandbox on a machine that has `uv` and a project virtualenv. Use it to run python:

```bash
# from the workspace root
cd repo
uv run --no-project --with python-pptx python build_deck.py     # build (write this script)
uv run --no-project --with python-pptx python main.py ttwo_overview.pptx   # validate
```

If `uv` is unavailable, fall back to `python3 -m pip install --user python-pptx` then plain `python3`.

`main.py ttwo_overview.pptx` prints a JSON report. **`"ok": true` with `score == max_score` (42/42) means success.** If `ok` is false, read each problem's `criterion`, `slide`, `shapes`, and `boxes_in`, fix exactly those named shapes in `build_deck.py`, rebuild, and re-validate. Repeat until `ok: true`.

You also have a `validate_deck` tool available that calls `check_pptx` directly on a path and returns the same JSON — use it as a faster alternative to the shell command.

## Definition of done

- `repo/ttwo_overview.pptx` exists and contains 14 slides matching the spec content.
- `check_pptx` returns `ok: true` (42/42).
- Summarize: the palette/motif you chose, how many fix-and-revalidate iterations it took, and the final score.
