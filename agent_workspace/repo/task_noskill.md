# Task: Generate a PowerPoint deck from the content spec

Build a polished `.pptx` slide deck from the content spec and validate its layout until it passes.

## Inputs (all paths relative to the workspace root)

- `repo/spec.md` — the content spec. It defines one slide per `##` section. Each slide has a full-sentence action title. Read it fully before writing any code — the spec is the source of truth for the deck's topic and slide count.
- `repo/harness/` + `repo/main.py` — the `check_pptx` layout validator. Run it on your output.
- `repo/reference_build_deck.py` — a **complete, working reference build script** for a *different* deck that already passes `check_pptx` with zero problems. Read it before you start. It shows how to use helper functions (`textbox`, `paragraph`, `rectangle`, `connector`, `style_run`, `_apply_bullet`, `title`, `chart`, `table`), how to set a palette and constants (margins, fixed title-top, source/page-number bands), and how to lay out cards, two-column bodies, charts-with-captions, and tables. **Reuse its helper functions and structural patterns verbatim**, then swap in YOUR palette and the content from `repo/spec.md`. Do NOT copy its Take-Two content — build the deck described in `repo/spec.md`. This reference is a guide, not the answer.

## What to produce

- The deck file you are told to build, built with **python-pptx** (16:9), one slide per `##` section in `spec.md`, in order.
- A bold, content-informed palette (a primary color + a darker variant for panels so white text keeps ≥4.5:1 contrast); dark cover/section/closing slides, light content slides; a repeated motif; every slide has a visual element; left-aligned body text; one title size deck-wide.

## Hard rules (these are what `check_pptx` enforces)

Build through reusable helpers (copy `textbox`, `paragraph`, `rectangle`, `connector`, `style_run`, `_apply_bullet`, and the higher-level `title`, `chart`, `table` from `reference_build_deck.py` into your `build_deck.py` and create every element through them):
- `textbox()` sets zero margins.
- `paragraph(..., bullet=True, level=n)` makes real `p.level` bullets (never a literal "• " typed into text).
- `title(slide, action_title, ...)` sizes the title box for its real wrap (no overflow), draws the accent connector beneath it, and returns the EMU y where content starts — use that return value to position content. NEVER hardcode a 0.7in title height.
- `chart(...)` turns OFF the default "Chart Title" and colors series from the palette.
- `table(...)` uses readable fonts (header ≥12pt, body ≥11pt) and real row heights.

`title()` imports `from harness.text_metrics import measure_text_emu` — run your build script from inside `repo/` (`cd repo`).

1. **Name every shape** with a meaningful `shape.name`.
2. **Bullets:** use `paragraph(..., bullet=True, level=n)` — real glyphs, never a literal "• " typed into the text.
3. **Full-bleed backgrounds** use `slide.background.fill` (the true `<p:bg>`), NOT a shape.
4. **No overlaps**, keep ≥0.5in from slide edges, align primary elements to one shared left edge across the deck, give row peers equal width + equal top, even gutters in rows/columns of 3+, **every chart and table needs a caption text box whose bottom is within 0.1in directly above it**, and keep the title at the same y on every slide.

## How to run things

```bash
cd repo
uv run --no-project --with python-pptx --with Pillow python build_deck.py        # build (write this script)
uv run --no-project --with python-pptx --with Pillow python main.py out.pptx      # validate
```

`main.py out.pptx` prints a JSON report. **`"ok": true` (zero problems) means success.** If `ok` is false, read each problem's `message`, `slide`, `shapes`, and `boxes_in`, fix exactly those named shapes in `build_deck.py`, rebuild, and re-validate. Repeat until `ok: true`. You also have a `validate_deck` tool that calls `check_pptx` directly on a path.

## Definition of done

- The deck file exists and contains one slide per `##` section in `spec.md`, matching the spec content.
- `check_pptx` returns `ok: true` (zero problems).
- Summarize: the palette/motif you chose, how many fix-and-revalidate iterations it took, and whether validation passed.
