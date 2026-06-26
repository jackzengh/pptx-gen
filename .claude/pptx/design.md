# Design Rules (mandatory)

These are concrete, non-negotiable rules that `check_pptx` cannot fully verify but a reader sees instantly. Follow all of them. They exist because real generated decks broke each one.

## Titles

- **Size the title box for the wrap, not for one line.** A 40pt title that wraps to 2 lines needs ~1.4in of height; 3 lines needs ~2.0in. If you give it 0.7in, the text overflows and **collides with the content below** — `check_pptx`'s `typography-text-fits-its-box` check WILL flag this. Use the `title()` helper, which sizes for up to 3 lines and pushes the subtitle/first content below the title's real bottom.
- **Every title gets a thin accent connector directly beneath it** (≈0.03in thick, ~1.6in wide, in the accent color), then the subtitle/content below that. This is the deck's spine motif. (This is the one place a line under a title is wanted — it's a deliberate brand element, distinct from the "no accent underline" anti-pattern, which refers to a full-width hairline faking structure.)
- Title y-position is the **same on every content slide** (`check_pptx` enforces `vertical-rhythm`). The connector and subtitle sit at fixed offsets below it.
- A chart/table **subtitle** (e.g. "Net revenue, FY ended March 31 ($M)") is a separate small caption that doubles as the element's required caption — place it directly above the chart/table, below the connector.

## Charts

- **NEVER leave the default "Chart Title".** python-pptx charts show a literal `Chart Title` placeholder unless you turn it off. Set `chart.has_title = False` and rely on the slide's subtitle/caption above the chart instead. The `chart()` helper does this.
- Style: set series colors from the palette, muted axis labels, hide the chart's own legend only if a single series; keep gridlines subtle.
- The chart still needs a **caption text box within 0.1in above it** (the subtitle) to satisfy `element-boxing`.

## Tables

- **Readable text:** header ≥ 12pt bold, body ≥ 11pt. Never let table text render tiny. Set explicit row heights (≥0.3in) so rows aren't cramped.
- Header row: filled with the primary/accent color, white bold text. Body rows: alternating subtle fill or none.
- The table needs a **caption within 0.1in above it** (its subtitle).

## Quote / section slides

- A pull-quote slide is large quote text + attribution. **Do not scatter loose colored bars in the middle of the slide** — orphaned motif elements read as a bug. If you use the chip/bar motif, anchor it consistently (e.g. top-right corner, same as content slides), not floating in empty space.
- Section dividers: big number or label on a dark background, one accent element, lots of breathing room. Keep it deliberate, not decorated.

## Stat cards / KPIs

- Cards in a row: equal width, equal top, even gutters (`check_pptx` enforces). Big number (≥28pt bold, accent color) + small label (≤12pt) beneath.
- Don't let a card's number+label overflow the card — size the card for the content.

## General

- Body text left-aligned; only titles/section labels centered if at all.
- Keep ≥0.5in from slide edges (hard fail inside 0.1in).
- One dominant color (60–70%), 1–2 supporting, one accent. Dark cover/closing, light content.
- Every text frame uses the `textbox()` helper (sets `auto_size = SHAPE_TO_FIT_TEXT`, zero margins). Every shape is named.

---

## Principles from professional / consulting decks

Distilled from McKinsey/BCG/Bain slide standards and presentation-design guides (see sources at bottom). These are the habits that separate a polished deck from an AI-looking one.

**Action titles (the single biggest lever).** A slide title is a *complete sentence stating the takeaway*, not a label. "Revenue grew 26% over three years to a record $3.4B" — not "Revenue Trend." Coined by BCG in the 1990s; it tells the reader what to conclude. The chart/table below is the *evidence* for that claim. Our spec already uses these — preserve them verbatim, don't shorten to topic labels.

**One idea per slide.** If a slide has two messages, split it. The action title is that one idea; everything on the slide supports it.

**Invisible grid + consistent structure.** Every slide shares the same margins, the same title position, the same body area, the same footnote line. Pick a left margin (~0.6in) and a title-top y, and use them on *every* content slide — the eye should never have to re-find the title. `check_pptx` enforces left-margin consistency and title vertical-rhythm precisely because consultants do this.

**Two-column is the workhorse layout.** Evidence on one side (chart/table/visual), takeaway bullets on the other. Dominates BCG/Bain decks. Vary it with stat-card rows and full-bleed section breaks so the deck isn't monotonous, but default to two-column for data slides.

**Highlight one number, grey the rest.** In a chart or table, emphasize the single data point that proves the title — bold color, a circle, or an arrow — and mute everything else to neutral grey. Don't rainbow-color every series. (Use the accent color for the hero point, supporting tones for context.)

**Whitespace is confidence.** Leave 15–20% of the slide empty. Generous margins and spacing improve comprehension and signal that the content is curated. Don't fill every inch; don't shrink text to cram more in. If it doesn't fit, cut content — never reduce body below 11pt or titles below 28pt.

**Typography = hierarchy through weight, not font count.** Max 2 font families (e.g. Georgia titles / Arial body, the McKinsey pairing). Create hierarchy with size and weight (bold/semibold/regular) within the family, not by adding fonts. Sans-serif reads best on screen for body. Test that the smallest text is legible at projection size.

**Source line on every data slide.** A small muted source/footnote at the bottom (e.g. "Source: FY2021 10-K"). Standard in consulting decks; we already require captions — keep the source line too.

**Charts clarify, not decorate.** Bar for comparison, line for trend over time, pie only for parts-of-a-whole with few slices. No chart junk: subtle gridlines, muted axis labels, no 3-D, no default "Chart Title", data labels only where they aid reading.
