---
name: deck-drafting
description: >
  Turn raw source material (a PDF, CSV/XLSX, report, or notes) into a content-only deck
  spec (Markdown) — the upstream step before building a .pptx. Use when the user wants a
  deck "from scratch" and hands you sources rather than a ready spec. This stage reads the
  sources, decides the deck narrative and what goes on each slide, prepares table and chart
  data, and writes a spec.md in this project's dialect. It does NOT write python-pptx or
  place anything on the slide — building, theming, layout, and check_pptx QA are the `pptx`
  skill's job.
---

# Deck Drafting (content spec from sources)

You are the **research / drafting** stage. Your job is to turn source material into a
**content-only deck spec** — a Markdown file describing the narrative, each slide's content,
and the data behind any table or chart. The separate **`pptx` skill** then reads your spec,
chooses a palette and layout, writes a `python-pptx` build script, and validates the result
with `check_pptx`. Palette, fonts, EMUs, and pixel placement are NOT your concern — only the
story, the content, and the data.

## What you own vs. what the `pptx` skill owns

| You (drafting) own                                            | The `pptx` skill owns                          |
| ------------------------------------------------------------- | ---------------------------------------------- |
| The deck narrative, slide `type`, slide-by-slide breakdown    | Palette, fonts, accent color                   |
| The content of each slide (bullets, cards, stats, steps)      | Exact geometry, alignment, gutters, margins    |
| The **data**: table values, chart categories + series, bins   | Writing the `python-pptx` `build_deck.py`      |
| Source citations and provenance                               | Fitting content; `check_pptx` QA; render verify |

Your `**Layout:**` line is an **advisory hint** in prose (e.g. "color sidebar left, hero image
right"), not coordinates. The `pptx` skill decides the real placement. Never emit grid cells,
EMUs, or pixel positions.

## How to work

1. **Read every source.** For a PDF, read it as a document. For CSV/XLSX, read the file and
   extract the actual numbers — never invent figures. Quote sources in each slide's `_Source:_`.
2. **Find the story.** What are the 5–10 things a reader must take away? Each becomes a slide
   with an **action title** — a full sentence stating the takeaway (not a label like "Revenue";
   instead "Net revenue grew 12% on strong King mobile bookings").
3. **Pick the slide type and the right component per point** (menus below). Every substantive
   slide needs a visual element: chart, table, number callout, process layout, or image. Avoid
   text-only slides except a sparse cover, section divider, quote, or closing.
4. **Prepare chart/table data explicitly.** For a histogram, compute the bins yourself and emit
   them as `**Chart (histogram):**` with categories = bin labels and one series of counts. For
   other charts, give the categories column + one column per series.
5. **Respect capacity heuristics** (below) so slides aren't overcrowded — if unsure, err toward
   fewer items per slide.
6. **Plan for diversity** (see _Diversity budget_). Assign each slide a distinct layout
   archetype before you write the body; never ship a deck where most slides share one skeleton.
7. **Write the spec** in the Markdown format below — content and data only, no palette.

## Diversity budget

Monotony is the most common failure mode. A deck that puts _title + rule + table/chart + source_
on nearly every slide reads as flat. To prevent that, **assign each slide a layout archetype as
you outline**, and enforce a budget across the deck:

- **No more than two consecutive slides** may share the same archetype.
- A deck of 6+ slides should use **at least four distinct archetypes**.
- **At least one** stat/number-callout slide and **at least one** non-table visual (chart, quote,
  or image) whenever the source data supports it.
- **Vary chart type** — don't make every chart a column chart.

Archetypes to rotate through (signal them via component choice + slide `type`; the `pptx` skill
lays them out):

| Archetype            | How you signal it in the spec                                                                                   |
| -------------------- | -------------------------------------------------------------------------------------------------------------- |
| **Stat callout**     | `type: data`/`content`; body = a compact `**Table:**` or `**Stats:**` of 1–4 headline numbers + short labels.  |
| **Single chart**     | `type: data`; body = one `**Chart (<type>):**`, varied chart type across the deck.                             |
| **Comparison table** | `type: comparison`; body = one `**Table:**` with a clear key column.                                           |
| **Side-by-side**     | `type: comparison`/`content`; two blocks meant to sit adjacent, e.g. a chart + a `**Bullets (right column):**`. |
| **Process / steps**  | `type: process`; body = 3–5 steps as a compact `**Table:**` or short `**Bullets:**`.                           |
| **Quote / callout**  | `type: quote`; body = one short `**Text:**` with the quote + attribution; no other blocks.                    |
| **Section divider**  | `type: section`; title only, optional one-line setup. Use to break a long deck into acts.                     |
| **Bulleted idea**    | `type: content`; body = one `**Bullets:**`. Use sparingly — it's the easiest to overuse.                       |

## Slide types

Set the `type` in each slide's `##` heading. It tells the `pptx` skill how the slide should feel;
it is not a coordinate or a PowerPoint master.

| type         | use for                          | drafting guidance                                      |
| ------------ | -------------------------------- | ------------------------------------------------------ |
| `cover`      | opening/title slide              | Deck title, subtitle or presenter line, no dense body. |
| `section`    | transition between topics        | Short title and optional one-line setup.               |
| `content`    | default idea slide               | Action title plus one or two components.               |
| `data`       | chart/table-first slide          | Clear figure title, explicit data, concise takeaway.   |
| `comparison` | compare options/segments         | Columns, table, or paired charts.                      |
| `process`    | sequence or workflow             | 3–5 concise steps as text or table rows.               |
| `quote`      | testimonial, excerpt, or callout | Quote text and attribution/source.                     |
| `closing`    | final takeaway or next steps     | Sparse summary, no new dense analysis.                 |

Match content type to layout: key points → bullets; team/category info often reads better as a
multi-column table; testimonials belong on quote slides; figures belong in charts or stat
callouts. Avoid repeating one text-heavy layout for every slide.

## Capacity heuristics (rough — the `pptx` skill enforces exactly via check_pptx)

- A content slide's body is one screen. Put **one or two** components on it, not five.
- Bullets: ≤ 5 bullets, ≤ ~18 words each.
- Number callouts: 1–4 headline figures, usually drafted as a compact table or stat cards.
- Process slides: 3–5 steps, usually a compact table or short bullets.
- Table: ≤ ~6 rows visible comfortably; more → split or summarize.
- Chart: ≤ ~8 categories for a bar/column; pie ≤ ~6 slices.

## Output format (Markdown)

Write the spec as a Markdown outline. The `pptx` skill reads it as text — there is no parser, so
write it for a human/model to read. **No palette, no coordinates.**

- One `# ` title line for the deck (optionally followed by an HTML `<!-- ... -->` note describing
  the intended style/density for the `pptx` skill).
- One `## ` heading per slide: `## <slide-id> — <type> — <action title>`, where `type` is one of
  `cover`, `section`, `content`, `data`, `comparison`, `process`, `quote`, `closing`.
- Optional `_Source: ..._` line under a slide for a citation.
- Body blocks as labeled lines/sub-sections. Use clear `**Label:**` prefixes the builder can read.
  Common labels in this project's specs:
  - `**Layout:** <prose>` — an advisory layout hint (e.g. "color sidebar left, hero image right").
  - `**Eyebrow:** <one line>` — a small kicker label above the title.
  - `**Text:** <one line>` — a single line (subtitle, callout, quote).
  - `**Body:** <prose>` — a paragraph of running text.
  - `**Bullets:**` (or `**Numbered list:**`) followed by a Markdown list (`- point` / `1. point`;
    prefix `**` for a bold lead-in).
  - `**Table:**` followed by a Markdown table (`| H1 | H2 |` / `|---|---|` / rows).
  - `**Chart (<chart_type>):** <chart title>` followed by a Markdown table whose first column is
    the categories and each remaining column is one series (header = series name). Put the exact
    numbers in the cells. `chart_type` ∈ `bar | column | line | line_markers | pie | area |
    histogram`. Add `**Chart subtitle:**` / `**Y-axis:**` lines when helpful.
  - `**Image:** <path or description>` — a local image path or an "Image note:" describing it.

You may use richer descriptive labels too (e.g. `**Stats (four hero cards):**`, `**Sidebar:**`,
`**Side table — …:**`) — the builder reads prose, so be specific about what each block is.

### Example

```markdown
# Activision Blizzard Q3 FY21

## cover — cover — Activision Blizzard Q3 FY21 Results
**Layout:** Color sidebar (left third) with white title; subtitle beneath.
**Text:** Q3 earnings summary · November 2021

## revenue — data — Net revenue grew across all three segments
_Source: ABK Q3 2021 10-Q_
**Chart (column):** Net revenue by segment ($M)
| Segment | Q3 2021 |
|---|---|
| Activision | 394 |
| Blizzard | 488 |
| King | 652 |

## margins — comparison — King leads on operating margin
**Table:**
| Segment | Revenue ($M) | Op. margin |
|---|---|---|
| Activision | 394 | 38% |
| King | 652 | 41% |

## closing — closing — Thank You
```

## Rules

- **Fully specific.** Every chart inlines its `chart_type`, title, categories, and exact series
  numbers; every table its headers and rows. No placeholders, no "TBD", no "data here" — the
  `pptx` skill adds ZERO data; it only renders what you wrote.
- **Never invent numbers.** Every figure must trace to a source you read.
- **No palette.** Do not specify colors or fonts — the `pptx` skill authors the theme.
- **No coordinates / no cells.** Content + data + slide type + advisory `**Layout:**` prose only.
- **Spend the diversity budget.** Rotate archetypes; no more than two consecutive slides share
  one; vary chart type. A near-uniform deck is a defect, not a style.
- **Action titles**, not labels.
- **Cite sources per slide** when using facts, quotes, or figures (the `_Source:_` line).
- Write the spec to `repo/spec.md` (the path the `pptx` skill reads) or to `specs/<deck-name>.md`,
  whichever the orchestrator/user asks for.
- After writing, briefly summarize the deck (slide count + one line each) so the narrative can be
  sanity-checked before building.

## Handoff to the `pptx` skill

Once the spec is written, the `pptx` skill reads it and: picks a task-matched palette, plans the
layout for each slide's blocks, writes `repo/build_deck.py` with `python-pptx`, runs it, and QAs
with `check_pptx` until the deck validates. Give it, per slide: the `type`, the action title, any
`_Source:_` line, an advisory `**Layout:**` hint, and fully-specified body blocks (with exact
chart/table data). It supplies everything visual.

If a slide feels too dense during drafting, condense it before handoff — one `##` section is one
slide, so split or shrink the content rather than expecting the builder to rewrite the story.
