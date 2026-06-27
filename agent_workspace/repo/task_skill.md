# Task: Generate a PowerPoint deck from the content spec

Build a polished `.pptx` slide deck from the content spec and validate its layout until it passes.

## Inputs (all paths relative to the workspace root)

- `repo/spec.md` — the content spec. It defines one slide per `##` section — typically a mix of stat-card KPI rows, line/column/pie charts, side-by-side tables, two-column bullet layouts, a section divider, a pull-quote, and a closing slide. Each slide has a full-sentence action title. Read it fully before writing any code — the spec is the source of truth for the deck's topic and slide count.
- `skills/pptx/SKILL.md` — the **PPTX skill**. Read it and the files it links (`design.md`, `python-pptx.md`, `naming.md`) before building. **Read `design.md` first** — it has the mandatory design rules AND a "Grader conventions" section (action-title verb lexicon, "Source:" not "Sources:", page numbers, one-size-per-level, <=6 hues, contrast >=4.5:1 with darker panels, fixed title-top, 0.3in margins). A separate rubric grader scores these; satisfy ALL of them. Follow it exactly.
- `repo/harness/` + `repo/main.py` — the `check_pptx` layout validator. Run it on your output.
- `repo/reference_build_deck.py` — a **complete, working reference build script** for a *different* deck that already passes `check_pptx` with zero problems. Read it before you start. It shows exactly how to copy in and use the skill's helpers (`textbox`, `paragraph`, `rectangle`, `connector`, `style_run`, `_apply_bullet`, `title`, `chart`, `table`), how to set the palette and constants (margins, fixed title-top, source/page-number bands), and how to lay out cards, two-column bodies, charts-with-captions, and tables that satisfy the grader. **Reuse its helper functions and structural patterns verbatim**, then swap in YOUR palette and the content from `repo/spec.md`. Do NOT copy its Take-Two content — build the deck described in `repo/spec.md`. This reference is a guide, not the answer.

## What to produce

- The deck file you are told to build (e.g. `repo/<deck>.pptx`) built with **python-pptx** (16:9), one slide per `##` section in `spec.md`, in order.
- Follow the skill's design guidance: a bold, content-informed palette (a primary color + a darker variant for panels so white text keeps ≥4.5:1 contrast); dark cover/section/closing slides, light content slides; a repeated motif; every slide has a visual element; left-aligned body text; one title size deck-wide per the grader's one-size-per-level rule.

## Hard rules from the skill (these are what `check_pptx` enforces)

**Build through the reusable helpers** in `skills/pptx/python-pptx.md` ("Reusable helpers" and "Higher-level helpers" sections) — copy `textbox`, `paragraph`, `rectangle`, `connector`, `style_run`, `_apply_bullet`, **and the higher-level `title`, `chart`, `table`** into your `build_deck.py` and create every element through them. This guarantees the rules below for free:
- `textbox()` sets zero margins.
- `paragraph(..., bullet=True, level=n)` makes real `p.level` bullets.
- **`title(slide, action_title, ...)`** sizes the title box for its real wrap (no overflow), draws the accent connector beneath it, and returns the EMU y where content starts — use that return value to position the chart/table/columns. NEVER hardcode a 0.7in title height.
- **`chart(...)`** turns OFF the default "Chart Title" and colors series from the palette. NEVER leave a chart showing "Chart Title".
- **`table(...)`** uses readable fonts (header ≥12pt, body ≥11pt) and real row heights.

`title()` imports `from harness.text_metrics import measure_text_emu` — that module is in `repo/harness/`, so run your build script from inside `repo/` (`cd repo`). Do NOT hand-roll `add_textbox`/`add_paragraph`/`add_chart` calls that skip these helpers.

1. **Name every shape** with a meaningful `shape.name` (e.g. `title`, `card_net_revenue`, `chart_revenue_trend`, `caption_chart_revenue`, `background_band`). See `skills/pptx/naming.md`.
2. **Bullets:** use `paragraph(..., bullet=True, level=n)` — real glyphs driven by `p.level`, never a literal "• " typed into the text.
3. **Full-bleed backgrounds** use `slide.background.fill` (the true `<p:bg>`), NOT a shape.
4. **No overlaps**, keep ≥0.5in from slide edges, align primary elements to one shared left edge across the deck, give row peers equal width + equal top, even gutters in rows/columns of 3+, and **every chart and table needs a caption text box whose bottom is within 0.1in directly above it**, and keep the title at the same y on every slide.

## How to run things

This workspace is a local sandbox on a machine that has `uv` and a project virtualenv. Use it to run python:

```bash
# from the workspace root — run from INSIDE repo/ so `import harness` resolves.
# Pillow is required: the validator's text-overflow check and the title() helper
# measure text with real fonts via PIL.
cd repo
uv run --no-project --with python-pptx --with Pillow python build_deck.py     # build (write this script)
uv run --no-project --with python-pptx --with Pillow python main.py <deck>.pptx   # validate (use the deck filename you were told to build)
```

If `uv` is unavailable, fall back to `python3 -m pip install --user python-pptx Pillow` then plain `python3`.

`main.py <deck>.pptx` prints a JSON report. **`"ok": true` (zero problems) means success.** If `ok` is false, read each problem's `message`, `slide`, `shapes`, and `boxes_in`, fix exactly those named shapes in `build_deck.py`, rebuild, and re-validate. Repeat until `ok: true`.

You also have a `validate_deck` tool available that calls `check_pptx` directly on a path and returns the same JSON — use it as a faster alternative to the shell command.

## Definition of done

- The deck file you were told to build exists and contains one slide per `##` section in `spec.md`, matching the spec content.
- `check_pptx` returns `ok: true` (zero problems).
- Summarize: the palette/motif you chose, how many fix-and-revalidate iterations it took, and whether validation passed.
